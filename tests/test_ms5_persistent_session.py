from __future__ import annotations

import json
from pathlib import Path

import pytest

from progress_studio.domain.mapping_models import ActivityRow, AllocationRecord, BOQRow
from progress_studio.infrastructure.session import (
    MappingSessionRepository,
    RecentSessionRepository,
    SessionValidationError,
)
from progress_studio.services.mapping_store import MappingStore


def make_store() -> MappingStore:
    store = MappingStore()
    store.load_activities([
        ActivityRow("A1000", "1", "1.1", "Floor demolition"),
        ActivityRow("A2000", "1", "1.2", "Wall demolition"),
    ])
    store.load_boq([
        BOQRow("Project|10|10", "Project", 10, "Demo", "Floor", "", "Remove floor", 1000.0),
        BOQRow("Project|11|11", "Project", 11, "Demo", "Wall", "", "Remove wall", 500.0),
    ])
    return store


def write_workbook_placeholder(path: Path, content: bytes) -> Path:
    path.write_bytes(content)
    return path


def test_session_round_trip_and_workbook_validation(tmp_path: Path) -> None:
    progress = write_workbook_placeholder(tmp_path / "progress.xlsx", b"progress-v1")
    boq = write_workbook_placeholder(tmp_path / "boq.xlsx", b"boq-v1")
    target = tmp_path / "project.mapping.json"
    repository = MappingSessionRepository()

    session = repository.create(
        progress,
        boq,
        "Project",
        [AllocationRecord("Project|10|10", "A1000", 35.0)],
    )
    repository.save(target, session)
    loaded = repository.load(target)

    assert loaded.boq_sheet == "Project"
    assert loaded.allocations == (
        AllocationRecord("Project|10|10", "A1000", 35.0),
    )
    assert repository.validate_workbook(loaded.progress) == progress.resolve()
    assert repository.validate_workbook(loaded.boq) == boq.resolve()


def test_session_write_is_json_and_has_expected_contract(tmp_path: Path) -> None:
    progress = write_workbook_placeholder(tmp_path / "progress.xlsx", b"progress")
    boq = write_workbook_placeholder(tmp_path / "boq.xlsx", b"boq")
    target = tmp_path / "project.mapping.json"
    repository = MappingSessionRepository()
    repository.save(target, repository.create(progress, boq, "NKC2", []))

    payload = json.loads(target.read_text(encoding="utf-8"))
    assert payload["format"] == "progress-studio-mapping-session"
    assert payload["version"] == 2
    assert payload["boq_sheet"] == "NKC2"
    assert payload["allocations"] == []
    assert not list(tmp_path.glob("*.tmp"))


def test_changed_workbook_is_rejected(tmp_path: Path) -> None:
    progress = write_workbook_placeholder(tmp_path / "progress.xlsx", b"progress-v1")
    boq = write_workbook_placeholder(tmp_path / "boq.xlsx", b"boq-v1")
    repository = MappingSessionRepository()
    session = repository.create(progress, boq, "Project", [])

    progress.write_bytes(b"progress-v2")
    with pytest.raises(SessionValidationError, match="does not match"):
        repository.validate_workbook(session.progress)


def test_store_restores_allocations_and_clears_undo_history() -> None:
    store = make_store()
    change = store.restore_allocations([
        AllocationRecord("Project|10|10", "A1000", 40.0),
        AllocationRecord("Project|10|10", "A2000", 60.0),
        AllocationRecord("Project|11|11", "A2000", 100.0),
    ])

    assert set(change.boq_keys) == {"Project|10|10", "Project|11|11"}
    assert store.activity_amount("A1000") == 400.0
    assert store.activity_amount("A2000") == 1100.0
    assert store.undo() is None


def test_store_rejects_invalid_session_records() -> None:
    store = make_store()
    with pytest.raises(ValueError, match="Activity was not found"):
        store.restore_allocations([
            AllocationRecord("Project|10|10", "MISSING", 50.0),
        ])
    with pytest.raises(ValueError, match="exceeds 100%"):
        store.restore_allocations([
            AllocationRecord("Project|10|10", "A1000", 60.0),
            AllocationRecord("Project|10|10", "A2000", 60.0),
        ])


def test_clear_all_is_one_undoable_command() -> None:
    store = make_store()
    store.restore_allocations([
        AllocationRecord("Project|10|10", "A1000", 100.0),
        AllocationRecord("Project|11|11", "A2000", 100.0),
    ])

    change = store.clear_all()
    assert store.mapped_amount == 0.0
    assert set(change.activity_ids) == {"A1000", "A2000"}

    store.undo()
    assert store.mapped_amount == 1500.0


def test_recent_sessions_keeps_latest_first_and_ignores_missing(tmp_path: Path) -> None:
    recent_file = tmp_path / "recent.json"
    repository = RecentSessionRepository(recent_file)
    first = tmp_path / "first.mapping.json"
    second = tmp_path / "second.mapping.json"
    first.write_text("{}", encoding="utf-8")
    second.write_text("{}", encoding="utf-8")

    repository.remember(first)
    repository.remember(second)
    repository.remember(first)

    assert repository.list() == [first.resolve(), second.resolve()]
    first.unlink()
    assert repository.list() == [second.resolve()]


def test_v1_session_is_migrated_to_current_version(tmp_path: Path) -> None:
    progress = write_workbook_placeholder(tmp_path / "progress.xlsx", b"progress")
    boq = write_workbook_placeholder(tmp_path / "boq.xlsx", b"boq")
    repository = MappingSessionRepository()
    current = repository.create(progress, boq, "Project", [])
    payload = {
        "format": current.format,
        "version": 1,
        "saved_at": current.saved_at,
        "progress": {
            "path": current.progress.path,
            "size": current.progress.size,
            "modified_ns": current.progress.modified_ns,
            "sha256": current.progress.sha256,
        },
        "boq": {
            "path": current.boq.path,
            "size": current.boq.size,
            "modified_ns": current.boq.modified_ns,
            "sha256": current.boq.sha256,
        },
        "boq_sheet": "Project",
        "allocations": [],
    }
    target = tmp_path / "legacy.mapping.json"
    target.write_text(json.dumps(payload), encoding="utf-8")

    loaded = repository.load(target)

    assert loaded.version == 2
    assert loaded.progress.filename == "progress.xlsx"
    assert loaded.boq.filename == "boq.xlsx"


def test_moved_or_renamed_workbook_can_be_relinked_by_hash(tmp_path: Path) -> None:
    original = write_workbook_placeholder(tmp_path / "progress.xlsx", b"same-content")
    repository = MappingSessionRepository()
    saved = repository.create(original, original, "Project", []).progress
    moved = tmp_path / "archive" / "renamed-progress.xlsx"
    moved.parent.mkdir()
    original.replace(moved)

    with pytest.raises(SessionValidationError, match="Browse for the moved"):
        repository.validate_workbook(saved)
    assert repository.validate_workbook(saved, moved) == moved.resolve()


def test_relink_rejects_different_workbook_content(tmp_path: Path) -> None:
    original = write_workbook_placeholder(tmp_path / "progress.xlsx", b"original")
    replacement = write_workbook_placeholder(tmp_path / "replacement.xlsx", b"different")
    repository = MappingSessionRepository()
    saved = repository.create(original, original, "Project", []).progress

    with pytest.raises(SessionValidationError, match="does not match"):
        repository.validate_workbook(saved, replacement)


def test_future_session_version_is_rejected_with_clear_message(tmp_path: Path) -> None:
    target = tmp_path / "future.mapping.json"
    target.write_text(
        json.dumps({"format": "progress-studio-mapping-session", "version": 999}),
        encoding="utf-8",
    )
    with pytest.raises(SessionValidationError, match="supports up to version"):
        MappingSessionRepository().load(target)
