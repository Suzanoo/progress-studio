from dataclasses import replace
from pathlib import Path
from unittest.mock import Mock

import pytest

from tests.fixtures.boq_revision import revision_case, edit, cache_formula
from progress_studio.infrastructure.session.mapping_session_repository import materialize_workbook_snapshot
from progress_studio.infrastructure.excel.mapping_reader import BOQSheetReader
from progress_studio.services.mapping_store import MappingStore


@pytest.fixture
def case(tmp_path, monkeypatch):
    c=revision_case(tmp_path)
    monkeypatch.setattr(c.repo,'materialize_snapshot',lambda snap:materialize_workbook_snapshot(snap,tmp_path/'cache'))
    return c


def preview(c):
    return c.service.preview(c.session,c.store,c.project,c.new)


@pytest.mark.smoke
def test_update_preserves_allocations_and_reopens_with_new_snapshot(case, tmp_path):
    c=case;original=c.project.read_bytes();state=dict(c.store.allocations)
    p=preview(c)
    assert len(p.item_changes)==2 and len(p.activity_changes)==2
    assert (p.before_total,p.after_total)==(3500,4100)
    assert (p.before_mapped,p.after_mapped)==(2900,3400)
    assert p.mapped_count==2
    assert c.project.read_bytes()==original and c.store.allocations==state
    staged,backup,cache=c.service.apply(p,c.session,c.store)
    assert c.store.boq_by_id['Project|2|2'].amount==1000
    assert staged.allocations==state
    assert staged.activity_amount('A1')==900 and staged.activity_amount('A2')==2500
    assert c.repo.load(backup).allocations==c.session.allocations
    restored=c.repo.load(c.project)
    assert restored.boq_snapshot.content_b64==p.updated.boq_snapshot.content_b64
    assert restored.progress_snapshot==c.session.progress_snapshot
    assert restored.working_tree_nodes==c.session.working_tree_nodes
    c.new.unlink();c.old.unlink()
    rows=BOQSheetReader().read(cache,'Project')
    fresh=MappingStore();fresh.load_activities(list(c.store.activities_by_id.values()))
    fresh.restore_working_tree(restored.working_tree_nodes);fresh.load_boq(rows)
    fresh.restore_allocations(list(restored.allocations))
    assert fresh.activity_amount('A2')==2500
    assert restored.boq.sha256==p.candidate_digest


def test_cancel_and_noop_never_write(case):
    c=case;original=c.project.read_bytes();p=preview(c)
    assert c.project.read_bytes()==original # discard preview = cancellation
    c.new.write_bytes(c.old.read_bytes());p=preview(c)
    assert not p.item_changes
    c.repo.save=Mock(side_effect=AssertionError('no write expected'))
    assert c.service.apply(p,c.session,c.store)==(c.store,None,None)


@pytest.mark.parametrize('mutate',[
    lambda s:s.__setitem__('E2','Different item'),
    lambda s:s.__setitem__('F2','m3'),
    lambda s:s.__setitem__('G2',21),
    lambda s:s.__setitem__('B2','Other WBS'),
    lambda s:s.__setitem__('A2','9.1'),
    lambda s:s.insert_rows(2),
    lambda s:s.delete_rows(3),
    lambda s:s.move_range('A2:H2',rows=5),
    lambda s:s.__setitem__('E5','Changed zero row'),
    lambda s:s.__setitem__('E6','Changed heading'),
    lambda s:s.insert_cols(2),
    lambda s:s.__setitem__('I1','Amount'),
])
def test_structural_revisions_are_rejected_without_mutation(case,mutate):
    c=case;old=c.project.read_bytes();alloc=dict(c.store.allocations)
    edit(c.new,mutate)
    with pytest.raises(ValueError):preview(c)
    assert c.project.read_bytes()==old and c.store.allocations==alloc


@pytest.mark.parametrize('cell,value',[('H2',0),('H2',None),('H5',1),('H6',1),('H5',None),('H2',-1),('H2','bad'),('H2',True),('H2','#VALUE!')])
def test_ineligible_or_invalid_amounts_rejected(case,cell,value):
    edit(case.new,lambda s:s.__setitem__(cell,value))
    with pytest.raises(ValueError):preview(case)


def test_duplicate_identity_is_not_row_based_remapping(case):
    def duplicate(s):
        for col in range(1,8):s.cell(3,col).value=s.cell(2,col).value
    edit(case.new,duplicate)
    with pytest.raises(ValueError,match='ambiguous'):preview(case)


def test_saved_formula_results_are_visible_and_used(case):
    cache_formula(case.new,'H2','750*2',1500)
    p=preview(case)
    assert p.new_formula_count==1 and p.old_formula_count==0
    staged,_,_=case.service.apply(p,case.session,case.store)
    assert staged.activity_amount('A1')==900


@pytest.mark.parametrize('value,kind',[(None,None),('#DIV/0!','e'),('wrong','str'),(True,'b'),(-1,None)])
def test_invalid_formula_results_rejected(case,value,kind):
    cache_formula(case.new,'H2','1/0',value,kind)
    with pytest.raises((ValueError,TypeError)):preview(case)


def test_old_formula_cache_also_required(case):
    cache_formula(case.old,'H2','1+999',None)
    session=replace(case.session,boq=case.repo.fingerprint(case.old),boq_snapshot=case.repo.snapshot(case.old))
    with pytest.raises(ValueError,match='formula result'):
        case.service.preview(session,case.store,case.project,case.new)


def test_cached_empty_heading_is_not_a_missing_numeric_result(case):
    # Existing BOQ Studio heading formulas return an explicitly cached empty string.
    cache_formula(case.new,'H6','IF(A6="","","")','','str')
    assert preview(case).new_formula_count==1
    cache_formula(case.new,'H2','""','','str')
    with pytest.raises(ValueError,match='eligibility'):preview(case)


def test_missing_string_cache_is_rejected(case):
    cache_formula(case.new,'H6','""',None,'str')
    with pytest.raises(ValueError,match='missing formula result'):preview(case)


def test_non_amount_identity_formula_requires_saved_evidence(case):
    cache_formula(case.new,'G2','10*2',None)
    with pytest.raises(ValueError,match='identity formula'):preview(case)


def test_non_amount_formula_result_change_is_structural(case):
    cache_formula(case.old,'G2','Source!A1',20)
    cache_formula(case.new,'G2','Source!A1',21)
    session=replace(case.session,boq=case.repo.fingerprint(case.old),boq_snapshot=case.repo.snapshot(case.old))
    with pytest.raises(ValueError,match='Only Amount'):
        case.service.preview(session,case.store,case.project,case.new)


@pytest.mark.parametrize('which',['candidate','project','mapping','tree','loaded_boq'])
def test_changes_after_preview_block_apply(case,which):
    c=case;p=preview(c)
    if which=='candidate':edit(c.new,lambda s:s.__setitem__('H2',1600))
    if which=='project':c.project.write_bytes(c.project.read_bytes()+b'\n')
    if which=='mapping':c.store.allocations[('Project|2|2','A1')]=50
    if which=='tree':c.store.load_activities([])
    if which=='loaded_boq':c.store.boq_by_id['Project|2|2']=replace(c.store.boq_by_id['Project|2|2'],amount=999)
    before=c.project.read_bytes()
    with pytest.raises(ValueError):c.service.apply(p,c.session,c.store)
    assert c.project.read_bytes()==before


@pytest.mark.parametrize('phase',['backup','project','cache'])
def test_failed_persistence_keeps_live_and_saved_state(case,monkeypatch,phase):
    c=case;p=preview(c);old=c.project.read_bytes();alloc=dict(c.store.allocations)
    save=c.repo.save
    def fail(path,session):
        if (phase=='backup' and path!=c.project) or (phase=='project' and path==c.project):raise OSError('injected disk failure')
        return save(path,session)
    monkeypatch.setattr(c.repo,'save',fail)
    if phase=='cache':monkeypatch.setattr(c.repo,'materialize_snapshot',Mock(side_effect=OSError('cache failure')))
    with pytest.raises(OSError):c.service.apply(p,c.session,c.store)
    assert c.project.read_bytes()==old and c.store.allocations==alloc
    assert c.store.boq_by_id['Project|2|2'].amount==1000
    if phase=='project':
        backup=next(c.project.parent.glob('*.before-boq-update-*.progressstudio'))
        assert c.repo.load(backup).allocations==c.session.allocations


def test_atomic_replace_failure_preserves_original_file(case,monkeypatch):
    import progress_studio.infrastructure.session.mapping_session_repository as module
    c=case;p=preview(c);old=c.project.read_bytes();real=module.os.replace
    def fail(source,target):
        if Path(target)==c.project:raise OSError('replace blocked')
        real(source,target)
    monkeypatch.setattr(module.os,'replace',fail)
    with pytest.raises(OSError):c.service.apply(p,c.session,c.store)
    assert c.project.read_bytes()==old
    assert not list(c.project.parent.glob('.*.tmp'))


def test_legacy_requires_verified_original_and_snapshot_handles_missing_source(case):
    c=case;legacy=replace(c.session,boq_snapshot=None)
    p=c.service.preview(legacy,c.store,c.project,c.new)
    assert p.item_changes
    c.old.unlink()
    with pytest.raises(ValueError):c.service.preview(legacy,c.store,c.project,c.new)
    assert preview(c).item_changes


def test_corrupt_embedded_source_is_rejected(case):
    snap=replace(case.session.boq_snapshot,content_b64='corrupt')
    with pytest.raises(ValueError):case.service.preview(replace(case.session,boq_snapshot=snap),case.store,case.project,case.new)


def test_invalid_existing_allocation_rejected(case):
    case.store.allocations[('Project|2|2','A1')]=90
    session=replace(case.session,allocations=tuple(case.store.allocation_records()))
    with pytest.raises(ValueError,match='exceeds'):case.service.preview(session,case.store,case.project,case.new)


def test_reopen_rejects_unreviewed_formula_cache_and_restores_snapshot(case):
    c = case
    cache_formula(c.new, 'H2', '750*2', 1500)
    p = preview(c)
    _, _, cache = c.service.apply(p, c.session, c.store)
    restored = c.repo.load(c.project)
    cache_formula(c.new, 'H2', '750*2', 9999)
    with pytest.raises(ValueError):
        c.repo.validate_workbook(restored.boq, c.new)
    c.repo.validate_workbook(restored.boq, cache)
    assert BOQSheetReader().read(cache, 'Project')[0].amount == 1500
