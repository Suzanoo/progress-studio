from datetime import datetime
import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from progress_studio.infrastructure.excel.duration_workbook import calendar_days, populate_durations, layout, milestone_ids
from progress_studio.infrastructure.excel.main_dataset_workbook_adapter import main_dataset_from_workbook
from progress_studio.services.rebuild_service import WorkbookRebuildEngine


@pytest.mark.parametrize('start,finish,expected',[
    (datetime(2026,1,2),datetime(2026,1,5),4),
    (datetime(2024,2,28),datetime(2024,3,1),3),
    (datetime(2026,1,1,20),datetime(2026,1,1,8),1),
    (None,datetime(2026,1,1),None),(datetime(2026,1,1),None,None),
    ('invalid',datetime(2026,1,1),None),(True,2,None),
    (4,3,None),(float('nan'),3,None),(2,float('inf'),None),(0,3,None),(2958466,2958467,None),
])
def test_inclusive_calendar_semantics(start,finish,expected):
    assert calendar_days(start,finish)==expected


def test_milestone_not_amount_and_epoch():
    assert calendar_days(datetime(2024,2,28),datetime(2024,3,1),milestone=True)==0
    assert calendar_days(datetime(2024,2,28),datetime(2024,3,1),epoch=CALENDAR_MAC_1904)==3
    assert calendar_days(None,None,milestone=True) is None


@pytest.mark.parametrize('dialect',['p6','msp'])
def test_create_layout_identity_formulas_and_rebuilds(tmp_path,monkeypatch,dialect):
    from tests.integration.create_progress.test_amount_first_workflow import _create
    source=_create(tmp_path,monkeypatch,dialect)
    w=load_workbook(source);header,h=layout(w['main']);ds=main_dataset_from_workbook(w)
    assert h['plan duration']+1==h['plan start'] and h['actual duration']+1==h['actual start']
    a=ds.activities[0];r=a.row_number
    for name in ('main','main_monthly','Weekly Compact'):
        assert w[name].cell(header,h['plan duration']).value=='Plan Duration'
        assert w[name].cell(header,h['actual duration']).value=='Actual Duration'
    s=w['main'];s.cell(r,h['amount'],0)  # Zero weight is still an ordinary activity.
    s.cell(r,h['plan start'],datetime(2024,2,28));s.cell(r,h['plan finish'],datetime(2024,3,1))
    s.cell(r+1,h['actual start'],datetime(2024,2,29,18));s.cell(r+1,h['actual finish'],datetime(2024,3,2,9))
    for item in ds.activities:
        pr=item.row_number;ar=pr+1
        assert s.cell(pr,h['plan duration']).data_type=='f'
        assert s.cell(ar,h['actual duration']).data_type=='f'
        assert s.cell(ar,h['plan duration']).value is None
        assert s.cell(pr,h['actual duration']).value is None
        assert s.cell(pr,h['plan duration']).protection.locked
        assert s.cell(ar,h['actual duration']).protection.locked
        assert 'TODAY' not in s.cell(ar,h['actual duration']).value
    assert ds.activities[-1].activity_id.upper() in milestone_ids(w)
    assert ',0,""' in s.cell(ds.activities[-1].row_number,h['plan duration']).value
    assert 'INT(' in s.cell(r,h['plan duration']).value
    before={c.coordinate:c.value for row in s for c in row if c.value is not None}
    w.save(source);w.close()
    for mode in ('snapshot','live'):
        out=tmp_path/f'{dialect}-{mode}.xlsx';engine=WorkbookRebuildEngine()
        (engine.rebuild_progress if mode=='snapshot' else engine.rebuild_live_progress)(source,out)
        w=load_workbook(out)
        assert {c.coordinate:c.value for row in w['main'] for c in row if c.value is not None}==before
        if mode=='snapshot':
            for name in ('main_monthly','Weekly Compact'):
                assert w[name].cell(r,h['plan duration']).value==3
                assert w[name].cell(r+1,h['actual duration']).value==3
        else:
            assert w['main_monthly'].cell(r,h['plan duration']).data_type=='f'
            assert "'main'!" in w['Weekly Compact'].cell(r,h['plan duration']).value
        assert all(w['main_monthly'].cell(row.row_number+1,h['actual duration']).value is None for row in ds.rows if row.row_type.lower() in ('wbs','project summary') and row.pa=='P')
        w.close()


def test_mapping_preserves_paired_actual_dates(tmp_path):
    from tests.integration.mapping.test_mapping_overlay_export import _created_mapping_input
    from progress_studio.services.workbook_export_service import WorkbookExportService
    source,store=_created_mapping_input(tmp_path,False)
    w=load_workbook(source);_,h=layout(w['main']);r=main_dataset_from_workbook(w).activities[0].row_number
    w['main'].cell(r+1,h['actual start'],datetime(2026,3,6))
    w['main'].cell(r+1,h['actual finish'],datetime(2026,3,13))
    w.save(source);w.close()
    out=tmp_path/'mapped.xlsx';WorkbookExportService().export(source,out,store)
    w=load_workbook(out);_,h=layout(w['main']);r=main_dataset_from_workbook(w).activities[0].row_number
    assert w['main'].cell(r+1,h['actual start']).value==datetime(2026,3,6)
    assert w['main'].cell(r+1,h['actual finish']).value==datetime(2026,3,13)
    assert w['main'].cell(r+1,h['actual duration']).data_type=='f'
    w.close()


def test_legacy_rebuild_does_not_add_columns(tmp_path):
    from tests.integration.rebuild.test_progress_rebuild_engine import _full_rebuild_fixture
    source=_full_rebuild_fixture(tmp_path/'old.xlsx')
    w=load_workbook(source);assert layout(w['main']) is None;w.close()
    for mode in ('snapshot','live'):
        out=tmp_path/f'old-{mode}.xlsx';engine=WorkbookRebuildEngine()
        (engine.rebuild_progress if mode=='snapshot' else engine.rebuild_live_progress)(source,out)
        w=load_workbook(out);assert layout(w['main']) is None
        assert all(c.value!='Plan Duration' for c in w['main_monthly'][4]);w.close()


def test_partial_schema_fails_explicitly():
    w=Workbook();s=w.active;s.title='main'
    s.append(['Row Type','Description','Plan Duration'])
    with pytest.raises(ValueError,match='no automatic migration'):populate_durations(w)


def test_summary_span_incomplete_actual_and_warning_idempotence():
    w=Workbook();s=w.active;s.title='main'
    s.append(['Row Type','Description','P/A','Activity ID','Plan Duration','Plan Start','Plan Finish','Actual Duration','Actual Start','Actual Finish'])
    s.append(['WBS','Summary','P',None,None,datetime(2026,1,1),datetime(2026,1,10)])
    s.append([None,None,'A',None,None,None,None,None,datetime(2026,1,1),datetime(2026,1,10)])
    s.append(['Activity','Ordinary','P','A1',None,datetime(2026,1,2),datetime(2026,1,3)])
    s.append([None,None,'A',None,None,None,None,None,datetime(2026,1,2),None])
    s.append(['Activity','Milestone','P','M1',None,datetime(2026,1,4),datetime(2026,1,4)])
    s.append([None,None,'A',None,None,None,None,None,datetime(2026,1,4),datetime(2026,1,4)])
    w.create_sheet('Info').append(['Calendar Duration Milestone: M1',True])
    populate_durations(w);count=len(s.conditional_formatting);populate_durations(w)
    assert len(s.conditional_formatting)==count
    out=w.copy_worksheet(s);populate_durations(w,out,snapshot=True)
    assert out['E2'].value==10 and out['H3'].value is None
    assert out['E4'].value==2 and out['H5'].value is None
    assert out['E6'].value==0 and out['H7'].value==0
    assert 'F4' in s['E4'].value and 'G4' in s['E4'].value
    assert any('NOT(AND(ISNUMBER(F4)' in r.formula[0] for area in s.conditional_formatting for r in s.conditional_formatting[area])
