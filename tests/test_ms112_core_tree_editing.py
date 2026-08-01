from pathlib import Path

import pytest

from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.infrastructure.session.mapping_session_repository import MappingSessionRepository
from progress_studio.services.mapping_store import MappingStore


def activity(activity_id: str, path, name: str) -> ActivityRow:
    return ActivityRow(
        activity_id=activity_id,
        parent_wbs=path[-1][0],
        child_wbs=path[-1][0],
        description=name,
        wbs_path=tuple(path),
    )


def make_store() -> MappingStore:
    store = MappingStore()
    store.load_activities([
        activity("A1000", (("1", "Project"), ("1.1", "Structure")), "Concrete"),
        activity("A1010", (("1", "Project"), ("1.1", "Structure")), "Steel"),
        activity("A2000", (("1", "Project"), ("1.2", "Architecture")), "Wall"),
    ])
    store.load_boq([
        BOQRow("B1", "BOQ", 2, "ARCH", "Wall", "", "Wall work", 100.0),
    ])
    return store


def test_original_wbs_can_be_renamed_and_undone() -> None:
    store = make_store()
    store.select_wbs((("1", "Project"), ("1.1", "Structure")))
    store.edit_selected_node(code="1.1", name="Structural Works")
    assert store.activities_by_id["A1000"].wbs_path[-1] == ("1.1", "Structural Works")
    assert store.undo_tree_edit()
    assert store.activities_by_id["A1000"].wbs_path[-1] == ("1.1", "Structure")
    assert store.redo_tree_edit()
    assert store.activities_by_id["A1010"].wbs_path[-1] == ("1.1", "Structural Works")


def test_original_activity_id_rename_keeps_allocations() -> None:
    store = make_store()
    store.toggle_activity("A1000")
    store.toggle_boq("B1")
    store.map_selected(100)
    store.edit_selected_node(code="A1005", name="Concrete Works")
    assert "A1000" not in store.activities_by_id
    assert store.activities_by_id["A1005"].description == "Concrete Works"
    assert store.allocation_share("B1", "A1005") == 100


def test_delete_rejects_mapped_activity_then_soft_deletes_after_unmap() -> None:
    store = make_store()
    store.toggle_activity("A1000")
    store.toggle_boq("B1")
    store.map_selected(100)
    with pytest.raises(ValueError, match="Unmap"):
        store.delete_selected_node()
    store.selected_boq_ids = {"B1"}
    store.unmap_selected()
    store.delete_selected_node()
    assert "A1000" not in store.activities_by_id
    assert store.undo_tree_edit()
    assert "A1000" in store.activities_by_id


def test_activity_can_be_reordered_and_reparented() -> None:
    store = make_store()
    store.toggle_activity("A1010")
    assert store.move_selected_node(-1)
    assert store.activity_order[:2] == ["A1010", "A1000"]
    destination = store.working_tree.find_wbs_by_path((("1", "Project"), ("1.2", "Architecture")))
    assert destination is not None
    store.reparent_selected_node(destination.node_id)
    assert store.activities_by_id["A1010"].wbs_path[-1] == ("1.2", "Architecture")


def test_session_v6_round_trips_working_tree(tmp_path: Path) -> None:
    store = make_store()
    store.select_wbs((("1", "Project"), ("1.1", "Structure")))
    store.edit_selected_node(code="1.1", name="Structure Revised")
    progress = tmp_path / "progress.xlsx"
    boq = tmp_path / "boq.xlsx"
    progress.write_bytes(b"progress")
    boq.write_bytes(b"boq")
    repo = MappingSessionRepository()
    session = repo.create(
        progress,
        boq,
        "Project",
        store.allocation_records(),
        store.supplemental_activities(),
        store.supplemental_wbs_nodes,
        list(store.working_tree_nodes()),
    )
    path = repo.save(tmp_path / "sample.progressstudio", session)
    loaded = repo.load(path)
    assert loaded.version == 6
    assert loaded.working_tree_nodes
    restored = make_store()
    restored.restore_working_tree(loaded.working_tree_nodes)
    assert restored.activities_by_id["A1000"].wbs_path[-1][1] == "Structure Revised"
