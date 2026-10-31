"""WU-1 placement safety and lifecycle, independent of native Excel acceptance."""
from copy import copy
from datetime import datetime
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from progress_studio.infrastructure.excel.cutoff_controls import (
    CutoffPlacementError, footer_control_row, find_cutoff_controls,
    remove_cell_validation, bound_cutoff_cell,
)
from progress_studio.infrastructure.excel.main_dataset_workbook_adapter import main_dataset_from_workbook
from progress_studio.infrastructure.excel.traditional_overlay_workbook import build_traditional_overlays
from progress_studio.infrastructure.excel.workbook_protection import apply_final_sheet_protection
from progress_studio.infrastructure.excel.weekly_compact_workbook import build_weekly_compact
from progress_studio.services.rebuild_service import WorkbookRebuildEngine


def _book(extra_columns=0):
    w=Workbook();s=w.active;s.title='main'
    headers=['Row Type','WBS','Description','P/A','Activity ID','Outline Level',
             'Plan Start','Plan Finish','Amount'] + [f'Extra {i}' for i in range(extra_columns)]
    for col,value in enumerate(headers,1):s.cell(4,col,value)
    for col,value in enumerate(['Activity','1','Cutoff Date','P','A1',1,
                                datetime(2026,1,3),datetime(2026,1,16),100],1):s.cell(5,col,value)
    s['D6']='A'
    for row,pa in enumerate(('P','AP','A','AA'),10):
        s.cell(row,1,'S-Curve');s.cell(row,3,'Plan' if pa=='P' else pa);s.cell(row,4,pa)
    for index,day in enumerate((9,16),len(headers)+1):
        s.cell(3,index,f'W{index-len(headers)}');s.cell(4,index,datetime(2026,1,day))
        s.cell(5,index,.5);s.cell(6,index,.25)
    w.copy_worksheet(s).title='main_monthly'
    d=w.create_sheet('Dashboard');d['K5']=datetime(2026,1,9)
    data=w.create_sheet('Dashboard_Data')
    for row,day in enumerate((9,16),2):
        for col,value in ((1,datetime(2026,1,day)),(2,.5*(row-1)),(3,.25*(row-1)),(10,datetime(2026,1,day)),
                          (4,datetime(2026,row-1,28)),(5,.5*(row-1)),(6,.25*(row-1)),(11,datetime(2026,row-1,28))):data.cell(row,col,value)
    return w,main_dataset_from_workbook(w)


@pytest.mark.parametrize('extra_columns',[0,2,7])
@pytest.mark.parametrize('legacy',['none','lm','description','duplicate'])
def test_old_new_layouts_rebuild_without_drift_and_preserve_unrelated_validation(tmp_path,extra_columns,legacy):
    w,ds=_book(extra_columns)
    for name in ('main','main_monthly'):
        s=w[name]
        if legacy in ('lm','duplicate'):
            s['L9']='Cutoff Date';s['M9']=datetime(2026,1,9)
            dv=DataValidation(type='list',formula1='"a,b"');dv.add('M8:N10');s.add_data_validation(dv)
        if legacy in ('description','duplicate'):
            s['C9']='Cutoff Date';s['I9']=datetime(2026,1,9)
            dv=DataValidation(type='list',formula1='"c,d"');dv.add('I9');s.add_data_validation(dv)
        s['F30']='keep note'
    for _ in range(3):
        build_traditional_overlays(w,main_dataset_from_workbook(w))
        apply_final_sheet_protection(w)
        for name,origin in [('main','J'),('main_monthly','K')]:
            s=w[name];c=bound_cutoff_cell(s)
            assert c.coordinate=='C14' and not c.protection.locked
            assert s['B14'].value=='Cutoff Date' and s['B14'].protection.locked
            assert s['C5'].value=='Cutoff Date' and s['F30'].value=='keep note'
            assert s['L9'].value is None and s['M9'].value is None
            assert s['C9'].value is None and s['I9'].value is None
            target=[v for v in s.data_validations.dataValidation if 'C14' in v.cells]
            assert len(target)==1 and f'${origin}$2' in target[0].formula1
            assert find_cutoff_controls(s,ds)==[(14,2,3)]
            if legacy in ('lm','duplicate'):
                assert any('N9' in v.cells for v in s.data_validations.dataValidation)
                assert all('M9' not in v.cells for v in s.data_validations.dataValidation)
        assert w['Dashboard']['K5'].value==datetime(2026,1,9)
        assert w['main']['C14'].number_format=='dd/mm/yyyy'
        assert w['main_monthly']['C14'].number_format=='mmm yyyy'
    out=tmp_path/'roundtrip.xlsx';w.save(out);w.close()
    with ZipFile(out) as archive:assert archive.testzip() is None
    w=load_workbook(out)
    assert next(w.defined_names['PS_WEEKLY_OVERLAY_CUTOFF'].destinations)==('main','$C$14')
    assert next(w.defined_names['PS_MONTHLY_OVERLAY_CUTOFF'].destinations)==('main_monthly','$C$14')
    w.close()


@pytest.mark.parametrize('initialization',['xml_field','manual_placeholder'])
def test_mapping_regeneration_retains_footer_controls(tmp_path,monkeypatch,initialization):
    from tests.integration.create_progress.test_ev_monetary_sources import test_mapping_regeneration_preserves_independent_bac_and_zero_milestone
    test_mapping_regeneration_preserves_independent_bac_and_zero_milestone(tmp_path,monkeypatch,initialization)
    w=load_workbook(tmp_path/'mapped.xlsx');row=footer_control_row(main_dataset_from_workbook(w))
    for name in ('main','main_monthly'):
        assert bound_cutoff_cell(w[name]).coordinate==f'C{row}'
        assert not w[name].cell(row,3).protection.locked
    assert w['Weekly Compact'].cell(row,3).value=='=PS_WEEKLY_OVERLAY_CUTOFF'
    w.close()


@pytest.mark.parametrize('conflict',['value','formula','comment','merged','validation','hyperlink'])
def test_footer_conflict_is_explicit_without_removing_old_control(conflict):
    w,ds=_book();s=w['main'];s['L9']='Cutoff Date';s['M9']=datetime(2026,1,9)
    if conflict=='value':s['C14']='User note'
    elif conflict=='formula':s['C14']='=1+2'
    elif conflict=='comment':s['C14'].comment=Comment('User note','User')
    elif conflict=='merged':s.merge_cells('B14:C14')
    elif conflict=='hyperlink':s['C14'].hyperlink='https://example.com'
    else:
        dv=DataValidation(type='list',formula1='"mine"');dv.add('C14');s.add_data_validation(dv)
    with pytest.raises(CutoffPlacementError,match='main.*[BC]14'):
        build_traditional_overlays(w,ds)
    assert s['M9'].value==datetime(2026,1,9) and s['L9'].value=='Cutoff Date'


def test_validation_subtraction_preserves_all_other_cells():
    w=Workbook();s=w.active
    dv=DataValidation(type='list',formula1='"x,y"');dv.add('A1:C3');dv.add('E7');s.add_data_validation(dv)
    remove_cell_validation(s,'B2')
    assert 'B2' not in dv.cells and 'E7' in dv.cells
    assert all(f'{c}{r}' in dv.cells for c in 'ABC' for r in range(1,4) if f'{c}{r}'!='B2')


@pytest.mark.parametrize('reference',["'main'!$C$14:$D$14", "'main_monthly'!$C$14",'OFFSET(main!$C$14,0,0)'])
def test_malformed_name_does_not_unlock_guessed_cells(reference):
    w,ds=_book();w['main']['B14']='Cutoff Date';w['main']['C14']=datetime(2026,1,9)
    w.defined_names.add(DefinedName('PS_WEEKLY_OVERLAY_CUTOFF',attr_text=reference))
    apply_final_sheet_protection(w)
    assert w['main']['C14'].protection.locked


@pytest.fixture(scope='module')
def generated(tmp_path_factory):
    from scripts.generate_weekly_compact_final_acceptance import build_acceptance
    return build_acceptance(tmp_path_factory.mktemp('wu1')/'source.xlsx')


def _input_values(w):
    ds=main_dataset_from_workbook(w);s=w['main']
    return {s.cell(row,col).coordinate:s.cell(row,col).value
            for item in ds.activities for row in (item.row_number,item.row_number+1)
            for col in range(1,s.max_column+1)}


@pytest.mark.parametrize('mode',['snapshot','live'])
def test_repeated_rebuild_and_legacy_snapshot_cleanup(generated,tmp_path,mode):
    source=tmp_path/'legacy.xlsx';w=load_workbook(generated)
    ds=main_dataset_from_workbook(w);new_row=footer_control_row(ds)
    for name,key in [('main','PS_WEEKLY_OVERLAY_CUTOFF'),('main_monthly','PS_MONTHLY_OVERLAY_CUTOFF')]:
        s=w[name];oldrow=min(r.row_number for r in ds.rows if r.row_type.lower()=='s-curve')-1
        s.cell(oldrow,12,'Cutoff Date');s.cell(oldrow,13,s.cell(new_row,3).value)
        s.cell(oldrow,13)._style=copy(s.cell(new_row,3)._style)
        remove_cell_validation(s,f'C{new_row}')
        s.cell(new_row,2).value=None;s.cell(new_row,3).value=None
        s.cell(new_row,3).comment=None
        dv=DataValidation(type='list',formula1='"legacy"');dv.add(f'M{oldrow}');s.add_data_validation(dv)
        w.defined_names[key]=DefinedName(key,attr_text=f"'{name}'!$M${oldrow}")
    w.create_sheet('User Notes')['A1']='Preserve'
    before=_input_values(w);w.save(source);w.close()
    engine=WorkbookRebuildEngine();method=engine.rebuild_progress if mode=='snapshot' else engine.rebuild_live_progress
    for index in range(2):
        out=tmp_path/f'{mode}-{index}.xlsx';method(source,out)
        w=load_workbook(out)
        assert _input_values(w)==before and w['User Notes']['A1'].value=='Preserve'
        for name in ('main','main_monthly'):
            assert bound_cutoff_cell(w[name]).coordinate==f'C{new_row}'
            assert not w[name].cell(new_row,3).protection.locked
            assert w[name].cell(oldrow,12).value is None
            assert len([v for v in w[name].data_validations.dataValidation if f'C{new_row}' in v.cells])==1
        compact=w['Weekly Compact']
        assert compact.cell(new_row,3).value=='=PS_WEEKLY_OVERLAY_CUTOFF'
        assert compact.cell(new_row,3).protection.locked
        for col in (12,13):
            coordinate=compact.cell(oldrow,col).coordinate
            assert compact.cell(oldrow,col).value in (None, f'=IF(\'main\'!{coordinate}="","",\'main\'!{coordinate})')
        assert not compact.data_validations.dataValidation
        w.close();source=out


@pytest.mark.parametrize('mode',['snapshot','live'])
@pytest.mark.parametrize('sheet',['main','main_monthly'])
def test_conflicting_rebuild_preserves_source_destination_and_cleans_temp(generated,tmp_path,mode,sheet):
    source=tmp_path/'source.xlsx';w=load_workbook(generated)
    ds=main_dataset_from_workbook(w);row=footer_control_row(ds)
    # This is user content, not a recognized generated control.
    w[sheet].cell(row,2,'User note');w[sheet].cell(row,3,'Do not replace')
    w.save(source);w.close();before=source.read_bytes()
    out=tmp_path/'existing.xlsx';out.write_bytes(b'KEEP')
    engine=WorkbookRebuildEngine();method=engine.rebuild_progress if mode=='snapshot' else engine.rebuild_live_progress
    with pytest.raises(CutoffPlacementError,match='overwrite'):
        method(source,out)
    assert source.read_bytes()==before and out.read_bytes()==b'KEEP'
    assert not list(tmp_path.glob('.existing.*'))


@pytest.mark.parametrize('dialect',['p6','msp'])
def test_amount_first_create_has_footer_controls(tmp_path,monkeypatch,dialect):
    from tests.integration.create_progress.test_amount_first_workflow import _create
    path=_create(tmp_path,monkeypatch,dialect)
    w=load_workbook(path);row=footer_control_row(main_dataset_from_workbook(w))
    for name in ('main','main_monthly'):
        assert bound_cutoff_cell(w[name]).coordinate==f'C{row}'
        assert not w[name].cell(row,3).protection.locked
        assert len([v for v in w[name].data_validations.dataValidation if f'C{row}' in v.cells])==1
    assert w['Weekly Compact'].cell(row,3).value=='=PS_WEEKLY_OVERLAY_CUTOFF'
    assert w['Weekly Compact'].cell(row,3).protection.locked
    w.close()
