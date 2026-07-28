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


def test_boq_wbs_filters_are_indexed_and_cascading():
    store = MappingStore(boq_page_size=10)
    rows = [
        BOQRow("K1", "S", 2, "Structure", "Foundation", "", "Concrete", 100.0),
        BOQRow("K2", "S", 3, "Structure", "Roof", "", "Steel", 50.0),
        BOQRow("K3", "S", 4, "Architecture", "Wall", "", "Brick", 25.0),
    ]
    store.load_boq(rows)

    assert store.boq_wbs2_values() == ("Architecture", "Structure")
    assert store.boq_wbs3_values("Structure") == ("Foundation", "Roof")

    store.boq_wbs2 = "Structure"
    assert store.boq_page_data().ids == ("K1", "K2")

    store.boq_wbs3 = "Roof"
    assert store.boq_page_data().ids == ("K2",)

    store.boq_query = "steel"
    assert store.boq_page_data().ids == ("K2",)
