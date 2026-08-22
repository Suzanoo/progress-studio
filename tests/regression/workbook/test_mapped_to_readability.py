from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.services.mapping_store import MappingStore


def make_store() -> MappingStore:
    store = MappingStore()
    store.load_activities([
        ActivityRow("A1000", "1", "1.1", "One"),
        ActivityRow("A2000", "1", "1.2", "Two"),
        ActivityRow("A3000", "1", "1.3", "Three"),
    ])
    store.load_boq([
        BOQRow("B1", "Project", 1, "ARCH", "Floor", "", "Item", 100.0),
    ])
    return store


def test_compact_mapped_to_text_preserves_first_activity_and_count():
    store = make_store()
    for activity_id, share in (("A1000", 20), ("A2000", 30), ("A3000", 50)):
        store.toggle_activity(activity_id)
        store.toggle_boq("B1")
        store.map_selected(share)

    assert store.mapped_to_compact_text("B1") == "A1000 (20%) +2"
    assert store.mapped_to_text("B1") == "A1000 (20%), A2000 (30%), A3000 (50%)"


def test_compact_mapped_to_text_for_unmapped_item():
    store = make_store()
    assert store.mapped_to_compact_text("B1") == "—"
