from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.services.mapping_store import MappingStore


def test_store_paginates_without_rendering_all_rows():
    store = MappingStore(activity_page_size=2, boq_page_size=3)
    store.load_activities([
        ActivityRow(f"A{i}", "1", f"1.{i}", f"Activity {i}") for i in range(5)
    ])
    store.load_boq([
        BOQRow(f"K{i}", "S", i, "Structure", "Foundation", "", f"Item {i}", float(i + 1))
        for i in range(8)
    ])

    assert store.activity_page_data().ids == ("A0", "A1")
    assert store.boq_page_data().ids == ("K0", "K1", "K2")
    store.boq_page = 3
    assert store.boq_page_data().ids == ("K6", "K7")


def test_map_and_undo_updates_store_only():
    store = MappingStore()
    store.load_activities([ActivityRow("A1", "1", "1.1", "Foundation")])
    store.load_boq([
        BOQRow("K1", "S", 2, "Structure", "Foundation", "", "Concrete", 100.0),
        BOQRow("K2", "S", 3, "Structure", "Foundation", "", "Rebar", 50.0),
    ])
    store.toggle_activity("A1")
    store.toggle_boq("K1")
    store.toggle_boq("K2")

    store.map_selected()
    assert store.activity_amount("A1") == 150.0
    assert store.mapped_amount == 150.0

    assert store.undo() is True
    assert store.assignments == {}
    assert store.mapped_amount == 0.0
