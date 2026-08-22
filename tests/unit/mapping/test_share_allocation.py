import pytest

from progress_studio.domain.mapping_models import ActivityRow, BOQRow, MappingStatus
from progress_studio.services.mapping_store import MappingStore


def make_store() -> MappingStore:
    store = MappingStore()
    store.load_activities([
        ActivityRow("A1000", "1", "1.1", "Floor demolition"),
        ActivityRow("A2000", "1", "1.2", "Ceiling demolition"),
    ])
    store.load_boq([
        BOQRow("Project:10", "Project", 10, "Demolition", "Floor", "", "Remove finishes", 1000.0),
    ])
    return store


def select(store: MappingStore, activity_id: str, boq_key: str = "Project:10") -> None:
    store.toggle_activity(activity_id)
    store.toggle_boq(boq_key)


def test_one_boq_can_be_split_between_multiple_activities() -> None:
    store = make_store()
    select(store, "A1000")
    store.map_selected(40)
    select(store, "A2000")
    store.map_selected(60)

    assert store.activity_amount("A1000") == 400.0
    assert store.activity_amount("A2000") == 600.0
    assert store.boq_share_percent("Project:10") == 100.0
    assert store.boq_allocated_amount("Project:10") == 1000.0
    assert store.boq_remaining_amount("Project:10") == 0.0
    assert store.boq_status("Project:10") is MappingStatus.FULL
    assert store.mapped_to_text("Project:10") == "A1000 (40%), A2000 (60%)"


def test_partial_allocation_reports_remaining_amount() -> None:
    store = make_store()
    select(store, "A1000")
    store.map_selected(25)

    assert store.boq_status("Project:10") is MappingStatus.PARTIAL
    assert store.boq_allocated_amount("Project:10") == 250.0
    assert store.boq_remaining_amount("Project:10") == 750.0
    assert store.mapped_amount == 250.0
    assert store.remaining_amount == 750.0


def test_total_share_cannot_exceed_one_hundred_percent() -> None:
    store = make_store()
    select(store, "A1000")
    store.map_selected(70)
    select(store, "A2000")
    with pytest.raises(ValueError, match="only 30% remaining"):
        store.map_selected(40)

    assert store.boq_share_percent("Project:10") == 70.0
    assert store.activity_amount("A2000") == 0.0


def test_invalid_share_is_rejected() -> None:
    store = make_store()
    select(store, "A1000")
    with pytest.raises(ValueError, match="greater than 0"):
        store.map_selected(0)
    with pytest.raises(ValueError, match="number"):
        store.map_selected("abc")


def test_unmap_removes_only_selected_activity_share() -> None:
    store = make_store()
    select(store, "A1000")
    store.map_selected(40)
    select(store, "A2000")
    store.map_selected(60)

    # A2000 remains selected; remove only its pair.
    store.toggle_boq("Project:10")
    store.unmap_selected()
    assert store.allocation_share("Project:10", "A1000") == 40.0
    assert store.allocation_share("Project:10", "A2000") == 0.0
    assert store.boq_status("Project:10") is MappingStatus.PARTIAL

    store.undo()
    assert store.boq_status("Project:10") is MappingStatus.FULL


def test_boq_remaining_percent_tracks_unallocated_share() -> None:
    store = make_store()

    assert store.boq_remaining_percent("Project:10") == 100.0

    store.selected_activity_ids = {"A1000"}
    store.selected_boq_ids = {"Project:10"}
    store.map_selected(35)

    assert store.boq_remaining_percent("Project:10") == 65.0

    store.selected_activity_ids = {"A2000"}
    store.selected_boq_ids = {"Project:10"}
    store.map_selected(65)

    assert store.boq_remaining_percent("Project:10") == 0.0
