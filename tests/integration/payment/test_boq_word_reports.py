from dataclasses import replace
from decimal import Decimal as D
from hashlib import sha256

import pytest
from docx import Document

from progress_studio.domain.boq_manual_requirements import compose, ManualInput
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.mapping_models import AllocationRecord
from progress_studio.services.boq_payment_export import ExportReview
from progress_studio.services.boq_word_report import BOQReportCatalog, export_boq_reports
from progress_studio.services.payment_plan_export import ExportCancelled
from tests.unit.payment.test_boq_payment_projection import item, payment, run


def catalog(tmp_path, p):
    path = tmp_path/'payment.xlsx'
    path.write_bytes(b'reviewed')
    return BOQReportCatalog(ExportReview(p, ((path, sha256(path.read_bytes()).hexdigest()),), False))


def branches():
    group = replace(item('g', role='group'), wbs='1')
    a = replace(item('a'), wbs='1.1', description='งานพื้น')
    b = replace(item('b'), wbs='1.2')
    c = replace(item('c'), wbs='2')
    p = project((group, a, b, c), (AllocationRecord('a','A',100), AllocationRecord('b','B',100)),
                payment(a=(1,1,.9), b=(1,1,.9)), ('A','B'))
    return compose(p, [ManualInput('c', (D('.5'), None, D('-.1')), 10)])


@pytest.mark.parametrize('keys,count,basis', [
    (['a'],1,100), (['g'],2,200), (['a','g'],2,200), (['g','c'],3,300),
    (['__project_summary__'],3,300)])
def test_selection_and_revision(tmp_path, keys, count, basis):
    cat = catalog(tmp_path, branches())
    plan = cat.select(keys)
    assert (plan.count, plan.total.amount) == (count, basis)
    if 'a' in keys and 'g' in keys:
        assert plan.details[1].ancestor.key == 'g'
    path = tmp_path/'selection.json'
    cat.save(path,set(keys))
    assert BOQReportCatalog(cat.review).restore(path) == set(keys)
    other = BOQReportCatalog(replace(cat.review,sources=((tmp_path/'other', 'different'),)))
    with pytest.raises(ValueError):
        other.restore(path)


def test_manual_gap_partial_and_tables(tmp_path):
    cat = catalog(tmp_path, branches())
    plan = cat.select({'g','a','c'})
    assert [d.row.key for d in plan.details] == ['g','a','c']
    assert plan.total.values[2] == (D(250),None,None)
    assert plan.total.values[3] == (D(250),None,D(-30))
    assert plan.known.period == (D(250),D(0),D(-30))
    output = export_boq_reports(plan,tmp_path,'Reports')
    for path in output.iterdir():
        doc = Document(path)
        assert not doc.settings.element.xpath('.//w:documentProtection')
        assert len(doc.tables)==4
        for index,t in enumerate(doc.tables[1:]):
            assert t.rows[0]._tr.xpath('./w:trPr/w:tblHeader')
            assert [tuple(c.text for c in row.cells) for row in t.rows[1:]] == list(
                plan.lines(index, 'CUMULATIVE' in path.name))
    assert 'งานพื้น' in list(plan.lines(0,True))[1][0]
    assert list(plan.lines(1,True))[0][2:]==('—','—')
    assert list(plan.lines(1,False))[0][2:]==('0.00%','0.00')
    assert list(plan.lines(2,True))[0][2]=='90.00%'


@pytest.mark.parametrize('schedule,expected', [
    ((1,1,1),('100.00%','—','—')), ((1,.5,1),('100.00%','50.00%','100.00%')),
    ((.999999,1,1),('<100%','100.00%','—'))])
def test_completion(tmp_path,schedule,expected):
    plan=catalog(tmp_path,run(a=schedule,b=schedule)).select(['b'])
    assert tuple(next(plan.lines(i,True))[2] for i in range(3))==expected
    assert all(next(plan.lines(i,False))[2]!='—' for i in range(3))
    for i, shown in enumerate(expected):
        if shown == '—':
            assert all(line[2:] == ('—','—') for line in plan.lines(i,True))


@pytest.mark.parametrize('p,state,known', [
    (run(shares=(60,)), 'PartialMapping', '48.00'), (run(shares=()), 'Unmapped', 'Unavailable'),
    (run(b=(None,None,None)), 'IncompleteRequirement', '48.00'),
    (run(amount=D(0)), 'Mapped', '0.00'), (run(amount=None), 'Mapped', 'Unavailable')])
def test_coverage_states(tmp_path,p,state,known):
    plan=catalog(tmp_path,p).select(['b'])
    rows=list(plan.lines(2,True))
    assert state in rows[0][1] and rows[-1][-1]==known
    if plan.total.amount in (None,0):
        assert plan.coverage is None and rows[-1][-2]=='Unavailable'


@pytest.mark.parametrize('fault',['cancel','changed','collision','second','ack'])
def test_pair_safety(tmp_path,monkeypatch,fault):
    from progress_studio.infrastructure.word import boq_plan
    cat=catalog(tmp_path,branches())
    if fault=='ack':
        cat=BOQReportCatalog(replace(cat.review,revision_required=True))
    plan=cat.select({'g'})
    render=boq_plan.render
    stopped=False
    def inject(path, plan, *, cumulative, check_cancel):
        nonlocal stopped
        if cumulative:
            if fault=='changed':
                cat.review.sources[0][0].write_bytes(b'changed externally')
            elif fault=='second':
                raise OSError('disk full')
            elif fault=='collision':
                (tmp_path/'Output').mkdir()
                (tmp_path/'Output'/'keep').write_text('existing')
            elif fault=='cancel':
                stopped=True
        render(path,plan,cumulative=cumulative,check_cancel=check_cancel)
    monkeypatch.setattr(boq_plan,'render',inject)
    with pytest.raises((OSError, ValueError, ExportCancelled)):
        export_boq_reports(plan,tmp_path,'Output',cancelled=lambda:stopped)
    assert not list(tmp_path.glob('.payment-plan-*'))
    assert not list(tmp_path.rglob('*.docx'))
    if fault=='collision':
        assert (tmp_path/'Output'/'keep').read_text()=='existing'
