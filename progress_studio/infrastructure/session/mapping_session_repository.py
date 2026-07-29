from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import tempfile

from progress_studio.domain.mapping_models import AllocationRecord
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
        size=stat.st_size,
        modified_ns=stat.st_mtime_ns,
        sha256=digest.hexdigest(),
    )


class MappingSessionRepository:
    """Persist mapping allocations as a small, atomic JSON sidecar."""

    def create(
        self,
        progress_file: Path,
        boq_file: Path,
        boq_sheet: str,
        allocations: list[AllocationRecord],
    ) -> MappingSessionData:
        return MappingSessionData(
            progress=fingerprint(progress_file),
            boq=fingerprint(boq_file),
            boq_sheet=boq_sheet,
            allocations=tuple(allocations),
            saved_at=datetime.now(timezone.utc).isoformat(),
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
        }
        _atomic_json_write(path, payload)
        return path

    def load(self, path: Path) -> MappingSessionData:
        path = Path(path)
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError as exc:
            raise SessionValidationError(f"Mapping session was not found: {path}") from exc
        except (OSError, json.JSONDecodeError) as exc:
            raise SessionValidationError(f"Mapping session cannot be read: {path}") from exc

        if payload.get("format") != SESSION_FORMAT:
            raise SessionValidationError("This file is not a Progress Studio mapping session.")
        if payload.get("version") != SESSION_VERSION:
            raise SessionValidationError(
                f"Unsupported mapping session version: {payload.get('version')!r}."
            )
        try:
            progress = WorkbookFingerprint(**payload["progress"])
            boq = WorkbookFingerprint(**payload["boq"])
            allocations = tuple(AllocationRecord(**item) for item in payload.get("allocations", []))
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
        )

    @staticmethod
    def validate_workbook(saved: WorkbookFingerprint) -> Path:
        path = Path(saved.path)
        current = fingerprint(path)
        if current.sha256 != saved.sha256:
            raise SessionValidationError(
                f"Workbook has changed since the session was saved: {path.name}"
            )
        return path


class RecentSessionRepository:
    """Keep a compact list of recently opened mapping-session files."""

    MAX_ITEMS = 10

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or (Path.home() / ".progress_studio" / "recent_sessions.json")

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
