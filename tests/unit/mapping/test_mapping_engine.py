from progress_studio.domain.mapping_models import ActivityRow, BOQRow, MappingStatus
from progress_studio.services.mapping_store import MappingStore


def _store() -> MappingStore:
    store = MappingStore(activity_page_size=10, boq_page_size=10)
    store.load_activities([
        ActivityRow("A1000", "1", "1.1", "Activity A"),
        ActivityRow("A2000", "2", "2.1", "Activity B"),
    ])
    store.load_boq([
        BOQRow("Project:2", "Project", 2, "W2", "W3", "W4", "Item 1", 100.0),
        BOQRow("Project:3", "Project", 3, "W2", "W3", "W4", "Item 2", 250.0),
    ])
    return store


def test_map_selected_allocates_full_amount_and_updates_balances() -> None:
    store = _store()
    store.toggle_activity("A1000")
    store.toggle_boq("Project:2")
    store.toggle_boq("Project:3")
    change = store.map_selected()
    assert change.boq_keys == ("Project:2", "Project:3")
    assert store.activity_amount("A1000") == 350.0
    assert store.boq_status("Project:2") is MappingStatus.FULL


def test_mapping_same_pair_replaces_share_and_undo_restores_it() -> None:
    store = _store()
    store.toggle_activity("A1000")
    store.toggle_boq("Project:2")
    store.map_selected(40)
    store.toggle_boq("Project:2")
    store.map_selected(65)
    assert store.activity_amount("A1000") == 65.0
    store.undo()
    assert store.activity_amount("A1000") == 40.0


def test_unmap_and_undo_restore_exact_allocation() -> None:
    store = _store()
    store.toggle_activity("A1000")
    store.toggle_boq("Project:3")
    store.map_selected()
    store.toggle_boq("Project:3")
    change = store.unmap_selected()
    assert change.activity_ids == ("A1000",)
    assert store.boq_status("Project:3") is MappingStatus.UNMAPPED
    store.undo()
    assert store.allocations[("Project:3", "A1000")] == 100.0


def test_mapping_requires_one_activity_and_at_least_one_boq() -> None:
    store = _store()
    try:
        store.map_selected()
    except ValueError as exc:
        assert str(exc) == "Select exactly one Activity."
    else:
        raise AssertionError("Expected a validation error")

    store.toggle_activity("A1000")
    try:
        store.map_selected()
    except ValueError as exc:
        assert str(exc) == "Select one or more BOQ items."
    else:
        raise AssertionError("Expected a validation error")
