import pytest

from progress_studio.domain.mapping_models import ActivityRow, BOQRow
from progress_studio.domain.working_tree import WorkingNodeKind
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
    store.load_boq([BOQRow("B1", "BOQ", 2, "STR", "Concrete", "", "Concrete work", 100.0)])
    return store


def test_created_activity_walks_below_selected_parent_not_at_tree_bottom() -> None:
    store = make_store()
    parent_path = (("1", "Project"), ("1.1", "Structure"))
    store.add_supplemental_activity(
        parent_path=parent_path,
        wbs_code="1.1",
        wbs_name="Structure",
        activity_id="A1020",
        description="Formwork",
    )
    rows = [(depth, node.code) for depth, node in store.working_tree.walk()]
    assert rows.index((3, "A1020")) < rows.index((2, "1.2"))
    structure = store.working_tree.find_wbs_by_path(parent_path)
    assert structure is not None
    assert [node.code for node in store.working_tree.children(structure.node_id)] == ["A1000", "A1010", "A1020"]


def test_indent_and_outdent_match_boq_tree_behavior_and_keep_mapping() -> None:
    store = make_store()
    store.toggle_activity("A2000")
    store.toggle_boq("B1")
    store.map_selected(100)

    store.select_wbs((("1", "Project"), ("1.2", "Architecture")))
    store.indent_selected_node()
    assert store.activities_by_id["A2000"].wbs_path == (
        ("1", "Project"), ("1.1", "Structure"), ("1.2", "Architecture")
    )
    assert store.allocation_share("B1", "A2000") == 100

    store.outdent_selected_node()
    assert store.activities_by_id["A2000"].wbs_path[-1] == ("1.2", "Architecture")
    assert store.allocation_share("B1", "A2000") == 100


def test_indent_requires_previous_wbs_sibling() -> None:
    store = make_store()
    store.toggle_activity("A1010")
    with pytest.raises(ValueError, match="previous sibling"):
        store.indent_selected_node()


def test_recursive_walk_contains_each_active_node_once() -> None:
    store = make_store()
    nodes = [node for _depth, node in store.working_tree.walk()]
    assert len(nodes) == len({node.node_id for node in nodes})
    assert {node.kind for node in nodes} == {WorkingNodeKind.WBS, WorkingNodeKind.ACTIVITY}
