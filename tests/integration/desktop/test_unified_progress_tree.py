from progress_studio.domain.mapping_models import ActivityRow
from progress_studio.services.mapping_store import MappingStore


def _store() -> MappingStore:
    store = MappingStore()
    store.load_activities([
        ActivityRow(
            activity_id="A1000",
            parent_wbs="4.1",
            child_wbs="4.1.1",
            description="Existing work",
            wbs_path=(("4", "Interior"), ("4.1", "1st Floor"), ("4.1.1", "Wall")),
        )
    ])
    return store


def test_add_activity_under_original_wbs() -> None:
    store = _store()
    parent = (("4", "Interior"), ("4.1", "1st Floor"))
    created = store.add_supplemental_activity(
        parent_path=parent,
        wbs_code="4.1",
        wbs_name="1st Floor",
        activity_id="A2500",
        description="New Activity",
    )
    assert created.wbs_path == parent
    assert created.parent_wbs == "4.1"
    assert created.is_supplemental


def test_add_sub_wbs_under_original_then_under_recent_created_wbs() -> None:
    store = _store()
    original = (("4", "Interior"), ("4.1", "1st Floor"))
    first = store.add_supplemental_wbs(parent_path=original, code="4.1.9", name="New Scope")
    assert store.selected_wbs_path == first.path
    second = store.add_supplemental_wbs(parent_path=store.selected_wbs_path, code="4.1.9.1", name="Sub Scope")
    assert second.parent_path == first.path
    assert store.selected_wbs_path == second.path


def test_add_activity_under_recent_created_wbs() -> None:
    store = _store()
    original = (("4", "Interior"), ("4.1", "1st Floor"))
    node = store.add_supplemental_wbs(parent_path=original, code="4.1.9", name="New Scope")
    created = store.add_supplemental_activity(
        parent_path=node.path,
        wbs_code=node.code,
        wbs_name=node.name,
        activity_id="A2510",
        description="Child Activity",
    )
    assert created.wbs_path == node.path
    assert created.parent_wbs == node.code
