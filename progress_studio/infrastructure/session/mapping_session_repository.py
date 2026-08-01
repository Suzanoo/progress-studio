from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import tempfile
from typing import Any, Callable

from progress_studio.domain.mapping_models import ActivityRow, AllocationRecord, SupplementalWBS
from progress_studio.domain.working_tree import WorkingScheduleTree
from progress_studio.infrastructure.platform_paths import user_data_dir

from progress_studio.domain.mapping_session import (
    MappingSessionData,
    SESSION_FORMAT,
    SESSION_VERSION,
    WorkbookFingerprint,
)


class SessionValidationError(ValueError):
    """Raised when a saved mapping session cannot be safely restored."""


def _atomic_json_write(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent)
    )
    temp_path = Path(temp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_path, path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def fingerprint(path: Path) -> WorkbookFingerprint:
    path = Path(path).expanduser().resolve()
    if not path.is_file():
        raise SessionValidationError(f"Workbook was not found: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    stat = path.stat()
    return WorkbookFingerprint(
        path=str(path),
        filename=path.name,
        size=stat.st_size,
        modified_ns=stat.st_mtime_ns,
        sha256=digest.hexdigest(),
    )


def _migrate_v1_to_v2(payload: dict[str, Any]) -> dict[str, Any]:
    """Add explicit workbook filenames while preserving v1 compatibility."""
    migrated = dict(payload)
    for key in ("progress", "boq"):
        workbook = dict(migrated.get(key) or {})
        workbook.setdefault("filename", Path(str(workbook.get("path", ""))).name)
        migrated[key] = workbook
    migrated["version"] = 2
    return migrated


Migration = Callable[[dict[str, Any]], dict[str, Any]]
def _migrate_v2_to_v3(payload: dict[str, Any]) -> dict[str, Any]:
    migrated = dict(payload)
    migrated.setdefault("supplemental_activities", [])
    migrated["version"] = 3
    return migrated


def _migrate_v3_to_v4(payload: dict[str, Any]) -> dict[str, Any]:
    migrated = dict(payload)
    migrated.setdefault("supplemental_wbs", [])
    migrated["version"] = 4
    return migrated


def _migrate_v4_to_v5(payload: dict[str, Any]) -> dict[str, Any]:
    """Assign stable editor identities to previously created nodes."""
    migrated = dict(payload)
    activities = []
    for item in migrated.get("supplemental_activities", []):
        record = dict(item)
        path = "/".join(part[0] for part in record.get("wbs_path", []))
        record.setdefault(
            "node_id",
            WorkingScheduleTree.legacy_created_node_id(
                "activity", f"{record.get('activity_id', '')}:{path}"
            ),
        )
        activities.append(record)
    wbs_nodes = []
    for item in migrated.get("supplemental_wbs", []):
        record = dict(item)
        parent = "/".join(part[0] for part in record.get("parent_path", []))
        record.setdefault(
            "node_id",
            WorkingScheduleTree.legacy_created_node_id(
                "wbs", f"{parent}:{record.get('code', '')}"
            ),
        )
        wbs_nodes.append(record)
    migrated["supplemental_activities"] = activities
    migrated["supplemental_wbs"] = wbs_nodes
    migrated["version"] = 5
    return migrated


_MIGRATIONS: dict[int, Migration] = {
    1: _migrate_v1_to_v2,
    2: _migrate_v2_to_v3,
    3: _migrate_v3_to_v4,
    4: _migrate_v4_to_v5,
}


def _migrate_payload(payload: dict[str, Any]) -> dict[str, Any]:
    try:
        version = int(payload.get("version"))
    except (TypeError, ValueError) as exc:
        raise SessionValidationError("The mapping session version is missing or invalid.") from exc

    if version > SESSION_VERSION:
        raise SessionValidationError(
            f"Unsupported mapping session version: {version}. "
            f"This application supports up to version {SESSION_VERSION}."
        )

    migrated = dict(payload)
    while version < SESSION_VERSION:
        migration = _MIGRATIONS.get(version)
        if migration is None:
            raise SessionValidationError(
                f"No migration path is available from mapping session version {version}."
            )
        migrated = migration(migrated)
        try:
            next_version = int(migrated.get("version"))
        except (TypeError, ValueError) as exc:
            raise SessionValidationError("A mapping session migration produced an invalid version.") from exc
        if next_version <= version:
            raise SessionValidationError("A mapping session migration did not advance the version.")
        version = next_version
    return migrated


class MappingSessionRepository:
    """Persist mapping allocations as a small, atomic JSON sidecar."""

    def create(
        self,
        progress_file: Path,
        boq_file: Path,
        boq_sheet: str,
        allocations: list[AllocationRecord],
        supplemental_activities: list[ActivityRow] | None = None,
        supplemental_wbs: list[SupplementalWBS] | None = None,
    ) -> MappingSessionData:
        return MappingSessionData(
            progress=fingerprint(progress_file),
            boq=fingerprint(boq_file),
            boq_sheet=boq_sheet,
            allocations=tuple(allocations),
            saved_at=datetime.now(timezone.utc).isoformat(),
            supplemental_activities=tuple(supplemental_activities or ()),
            supplemental_wbs=tuple(supplemental_wbs or ()),
        )

    def save(self, path: Path, session: MappingSessionData) -> Path:
        path = Path(path)
        payload = {
            "format": session.format,
            "version": session.version,
            "saved_at": session.saved_at,
            "progress": asdict(session.progress),
            "boq": asdict(session.boq),
            "boq_sheet": session.boq_sheet,
            "allocations": [asdict(record) for record in session.allocations],
            "supplemental_activities": [asdict(record) for record in session.supplemental_activities],
            "supplemental_wbs": [asdict(record) for record in session.supplemental_wbs],
        }
        _atomic_json_write(path, payload)
        return path

    def load(self, path: Path) -> MappingSessionData:
        path = Path(path)
        try:
            raw_payload = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError as exc:
            raise SessionValidationError(f"Mapping session was not found: {path}") from exc
        except (OSError, json.JSONDecodeError) as exc:
            raise SessionValidationError(f"Mapping session cannot be read: {path}") from exc

        if not isinstance(raw_payload, dict):
            raise SessionValidationError("The mapping session root must be a JSON object.")
        if raw_payload.get("format") != SESSION_FORMAT:
            raise SessionValidationError("This file is not a Progress Studio mapping session.")
        payload = _migrate_payload(raw_payload)
        try:
            progress = WorkbookFingerprint(**payload["progress"])
            boq = WorkbookFingerprint(**payload["boq"])
            allocations = tuple(AllocationRecord(**item) for item in payload.get("allocations", []))
            supplemental_activities = tuple(
                ActivityRow(**{**item, "wbs_path": tuple(tuple(part) for part in item.get("wbs_path", ()))})
                for item in payload.get("supplemental_activities", [])
            )
            supplemental_wbs = tuple(
                SupplementalWBS(**{**item, "parent_path": tuple(tuple(part) for part in item.get("parent_path", ()))})
                for item in payload.get("supplemental_wbs", [])
            )
            boq_sheet = str(payload["boq_sheet"]).strip()
            saved_at = str(payload["saved_at"])
        except (KeyError, TypeError, ValueError) as exc:
            raise SessionValidationError("The mapping session is incomplete or malformed.") from exc
        if not boq_sheet:
            raise SessionValidationError("The mapping session does not specify a BOQ worksheet.")
        return MappingSessionData(
            progress=progress,
            boq=boq,
            boq_sheet=boq_sheet,
            allocations=allocations,
            saved_at=saved_at,
            supplemental_activities=supplemental_activities,
            supplemental_wbs=supplemental_wbs,
        )

    @staticmethod
    def validate_workbook(saved: WorkbookFingerprint, candidate: Path | None = None) -> Path:
        """Return a verified workbook path.

        The saved absolute path is tried by default. A user-selected candidate may be
        supplied when the workbook was moved or renamed. Only identical SHA-256 content
        is accepted; changed workbooks are never merged automatically.
        """
        path = Path(candidate).expanduser().resolve() if candidate else saved.saved_path
        try:
            current = fingerprint(path)
        except SessionValidationError as exc:
            if candidate is None:
                raise SessionValidationError(
                    f"{saved.filename} was not found at its saved location. "
                    "Browse for the moved or renamed workbook to continue."
                ) from exc
            raise
        if current.sha256 != saved.sha256:
            source = "selected workbook" if candidate else "workbook at the saved location"
            raise SessionValidationError(
                f"The {source} does not match the session copy: {saved.filename}. "
                "Progress Studio will not merge a changed workbook automatically."
            )
        return path


class RecentSessionRepository:
    """Keep a compact list of recently opened mapping-session files."""

    MAX_ITEMS = 10

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or (user_data_dir() / "recent_sessions.json")

    def list(self) -> list[Path]:
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, OSError, json.JSONDecodeError):
            return []
        result: list[Path] = []
        for raw in payload if isinstance(payload, list) else []:
            path = Path(str(raw))
            if path.is_file() and path not in result:
                result.append(path)
        return result[: self.MAX_ITEMS]

    def remember(self, session_path: Path) -> None:
        session_path = Path(session_path).expanduser().resolve()
        items = [path for path in self.list() if path != session_path]
        items.insert(0, session_path)
        _atomic_json_write(self.path, [str(path) for path in items[: self.MAX_ITEMS]])
