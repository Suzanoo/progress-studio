from progress_studio.domain.mapping_models import ActivityRow
from progress_studio.services.mapping_store import MappingStore


def test_add_and_remove_supplemental_activity():
    store = MappingStore()
    store.load_activities([ActivityRow('A1000', '', '1', 'Base', (('1', 'Project'),))])
    created = store.add_supplemental_activity(
        parent_path=(('1', 'Project'),),
        wbs_code='1.X1',
        wbs_name='Waterproofing',
        activity_id='PS-A001',
        description='Waterproofing Work',
    )
    assert created.is_supplemental
    assert created.wbs_path[-1] == ('1.X1', 'Waterproofing')
    assert store.supplemental_activities() == [created]
    store.remove_supplemental_activity('PS-A001')
    assert 'PS-A001' not in store.activities_by_id


def test_cannot_duplicate_supplemental_activity_id():
    store = MappingStore()
    store.load_activities([ActivityRow('A1000', '', '1', 'Base')])
    try:
        store.add_supplemental_activity(
            parent_path=(), wbs_code='PS-1', wbs_name='Extra',
            activity_id='A1000', description='Duplicate',
        )
    except ValueError as exc:
        assert 'already exists' in str(exc)
    else:
        raise AssertionError('duplicate Activity ID was accepted')
