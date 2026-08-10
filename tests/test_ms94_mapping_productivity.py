from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.services.mapping_store import MappingStore


def _boq(index: int, wbs2: str = "Structure") -> BOQRow:
    return BOQRow(
        key=f"B{index}",
        source_sheet="Project",
        source_row=index,
        wbs2=wbs2,
        wbs3="Foundation",
        wbs4="Concrete",
        description=f"BOQ item {index}",
        amount=float(index * 100),
    )


def _activity(activity_id: str, path: tuple[tuple[str, str], ...]) -> ActivityRow:
    return ActivityRow(
        activity_id=activity_id,
        parent_wbs=path[0][0],
        child_wbs=path[-1][0],
        description=activity_id,
        wbs_path=path,
    )


def test_boq_range_page_and_filtered_selection_persist_across_pages() -> None:
    store = MappingStore(boq_page_size=2)
    store.load_boq([_boq(i) for i in range(1, 6)])

    store.toggle_boq("B2")
    assert store.select_boq_range("B4") == ("B2", "B3", "B4")
    assert store.selected_boq_ids == {"B2", "B3", "B4"}

    store.boq_page = 3
    assert store.select_boq_page() == ("B5",)
    assert store.selected_boq_ids == {"B2", "B3", "B4", "B5"}

    store.boq_query = "item 1"
    assert store.select_all_filtered_boq() == ("B1",)
    assert store.selected_boq_ids == {"B1", "B2", "B3", "B4", "B5"}
    assert store.selected_boq_amount == 1500.0


def test_clear_boq_selection_resets_anchor() -> None:
    store = MappingStore()
    store.load_boq([_boq(1), _boq(2)])
    store.toggle_boq("B1")
    store.clear_boq_selection()
    assert store.selected_boq_ids == set()
    assert store.boq_selection_anchor is None


def test_wbs_collapse_retains_one_header_representative_and_expand_restores_rows() -> None:
    foundation = (("1", "Structure"), ("1.1", "Foundation"))
    frame = (("1", "Structure"), ("1.2", "Frame"))
    store = MappingStore(activity_page_size=20)
    store.load_activities([
        _activity("A1", foundation),
        _activity("A2", foundation),
        _activity("A3", frame),
    ])

    store.toggle_wbs(("1", "1.1"))
    assert store.activity_page_data().ids == ("A1", "A3")
    assert not store.is_activity_visible("A1")
    assert store.is_activity_visible("A3")

    store.expand_all_wbs()
    assert store.activity_page_data().ids == ("A1", "A2", "A3")


def test_search_ignores_collapsed_state_to_show_matches() -> None:
    path = (("1", "Structure"), ("1.1", "Foundation"))
    store = MappingStore()
    store.load_activities([_activity("A1", path), _activity("A2", path)])
    store.collapse_all_wbs()
    store.activity_query = "A2"
    assert store.activity_page_data().ids == ("A2",)
