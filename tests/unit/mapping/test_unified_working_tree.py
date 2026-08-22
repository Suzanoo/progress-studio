from __future__ import annotations

from progress_studio.domain.mapping_models import ActivityRow
from progress_studio.domain.working_tree import (
    WorkingNodeKind,
    WorkingNodeOrigin,
    WorkingScheduleTree,
)
from progress_studio.services.mapping_store import MappingStore


def _rows() -> list[ActivityRow]:
    return [
        ActivityRow(
            activity_id="A1000",
            parent_wbs="4.1.1",
            child_wbs="4.1.1",
            description="Existing wall work",
            wbs_path=(("4", "Interior"), ("4.1", "1st Floor"), ("4.1.1", "Wall")),
        ),
        ActivityRow(
            activity_id="A1010",
            parent_wbs="4.1.2",
            child_wbs="4.1.2",
            description="Existing ceiling work",
            wbs_path=(("4", "Interior"), ("4.1", "1st Floor"), ("4.1.2", "Ceiling")),
        ),
    ]


def test_workbook_nodes_have_stable_identity_across_rebuilds() -> None:
    first = WorkingScheduleTree.build(_rows())
    second = WorkingScheduleTree.build(_rows())

    assert first.find_activity("A1000").node_id == second.find_activity("A1000").node_id
    path = (("4", "Interior"), ("4.1", "1st Floor"))
    assert first.find_wbs_by_path(path).node_id == second.find_wbs_by_path(path).node_id


def test_store_exposes_one_tree_for_original_and_created_nodes() -> None:
    store = MappingStore()
    store.load_activities(_rows())
    parent = (("4", "Interior"), ("4.1", "1st Floor"))
    created_wbs = store.add_supplemental_wbs(
        parent_path=parent,
        code="4.1.9",
        name="Added During Mapping",
    )
    created_activity = store.add_supplemental_activity(
        parent_path=created_wbs.path,
        wbs_code=created_wbs.code,
        wbs_name=created_wbs.name,
        activity_id="A2500",
        description="Added Activity",
    )

    original = store.working_tree.find_activity("A1000")
    created = store.working_tree.find_activity("A2500")
    created_parent = store.working_tree.find_wbs_by_path(created_wbs.path)

    assert original is not None and original.origin is WorkingNodeOrigin.WORKBOOK
    assert created is not None and created.origin is WorkingNodeOrigin.USER_CREATED
    assert created.kind is WorkingNodeKind.ACTIVITY
    assert created.parent_id == created_parent.node_id
    assert created.node_id == created_activity.node_id
    assert store.working_tree.validate() == ()


def test_selected_node_identity_survives_tree_rebuild() -> None:
    store = MappingStore()
    store.load_activities(_rows())
    path = (("4", "Interior"), ("4.1", "1st Floor"))
    store.select_wbs(path)
    selected_before = store.selected_node_id

    store.add_supplemental_wbs(parent_path=path, code="4.1.9", name="New Scope")
    store.select_wbs(path)

    assert selected_before
    assert store.selected_node_id == selected_before


def test_created_node_ids_are_not_display_codes() -> None:
    store = MappingStore()
    store.load_activities(_rows())
    parent = (("4", "Interior"),)
    created = store.add_supplemental_wbs(parent_path=parent, code="4.9", name="New")
    node = store.working_tree.find_wbs_by_path(created.path)

    assert node is not None
    assert node.node_id == created.node_id
    assert node.node_id != created.code
