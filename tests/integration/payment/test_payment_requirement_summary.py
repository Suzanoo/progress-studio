"""PI-1 preservation, failure and calculated workbook behavior."""
from pathlib import Path
import shutil
import subprocess

import pytest
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
from progress_studio.infrastructure.excel.payment_input_reader import PaymentInputSparseReader
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError
from progress_studio.infrastructure.excel.workbook_protection import _protect_payment_input
from progress_studio.services.payment_service import PaymentService
from tests.fixtures.payment_requirements import requirement_source, edit_requirements
from tests.fixtures.boq_revision import cache_formula


def values(path):
    data=PaymentInputSparseReader().read(path)
    return {p.period_id:{r.activity_id:r.required_fraction for r in p.requirements} for p in data.periods}


def activity_rows(ws):
    return {ws.cell(r,3).value:r for r in range(8,ws.max_row+1) if ws.cell(r,1).value=='ACT'}


@pytest.mark.parametrize('embedded',[True,False])
def test_amount_hierarchy_protection_and_sparse_reconcile(tmp_path,embedded):
    source=requirement_source(tmp_path/'source.xlsx');output=tmp_path/'input.xlsx'
    if embedded:PaymentService().prepare_embedded_payment_input(source,output,3)
    else:PaymentInputWorkbook().create(source,output,3)
    edit_requirements(output)
    before=values(output)
    wb=load_workbook(output);ws=wb['Payment Input'];rows=activity_rows(ws)
    assert ws['E6'].value=='Amount' and ws['F6'].value=='P01' and ws.freeze_panes=='F8'
    assert ws.cell(rows['A100'],5).value==100
    assert ws.cell(rows['M1'],5).value==0
    assert ws.cell(rows['A400'],5).value=='Missing Amount'
    assert ws.cell(rows['A100'],5).number_format=='#,##0.00'
    assert ws.protection.sheet and wb.calculation.calcMode=='manual' and wb.calculation.calcOnSave
    if not embedded:
        assert all(not ws.cell(7,c).protection.locked for c in range(6,9))
    for r in range(8,ws.max_row+1):
        assert ws.cell(r,5).protection.locked
        for c in range(6,9):assert ws.cell(r,c).protection.locked == (ws.cell(r,1).value!='ACT')
    assert len([r for r in range(8,ws.max_row+1) if ws.cell(r,4).value=='Repeated branch'])==2
    assert before['P02']=={'A100':0,'M1':0} and before['P03']=={}
    if embedded:
        wb['main']['O7']=125.25
        wb.save(output);wb.close()
        PaymentService().prepare_embedded_payment_input(output,output,3)
        assert values(output)==before
        wb=load_workbook(output);ws=wb['Payment Input']
        assert ws.cell(activity_rows(ws)['A100'],5).value==125.25
        assert wb['User Notes']['A1'].value=='PRESERVE'
    wb.close()


@pytest.mark.parametrize('raw,expected',[(None,'Missing Amount'),(0,0),(-1,'Invalid Amount'),
    (True,'Invalid Amount'),('250','Invalid Amount'),('#VALUE!','Invalid Amount'),
    ('=100+25','Formula unavailable'),(12.345,12.345)])
@pytest.mark.parametrize('embedded',[True,False])
def test_amount_status_consistent(tmp_path,raw,expected,embedded):
    source=requirement_source(tmp_path/'source.xlsx')
    w=load_workbook(source);w['main']['O7']=raw;w.save(source);w.close()
    out=tmp_path/'out.xlsx'
    if embedded:PaymentService().prepare_embedded_payment_input(source,out,3)
    else:PaymentInputWorkbook().create(source,out,3)
    w=load_workbook(out);s=w['Payment Input']
    assert s.cell(activity_rows(s)['A100'],5).value==expected
    w.close()


@pytest.mark.parametrize('cached,kind,expected',[(125.5,None,125.5),(0,None,0),(None,None,'Formula unavailable'),
    (-2,None,'Invalid Amount'),('NaN',None,'Invalid Amount'),('1','b','Invalid Amount'),('#VALUE!','e','Invalid Amount')])
def test_saved_formula_results_matched_by_id(tmp_path,cached,kind,expected):
    source=requirement_source(tmp_path/'source.xlsx')
    cache_formula(source,'O7','100+25.5',cached,kind)
    for mode in ('standalone','embedded'):
        out=tmp_path/f'{mode}.xlsx'
        if mode=='embedded':PaymentService().prepare_embedded_payment_input(source,out,3)
        else:PaymentInputWorkbook().create(source,out,3)
        w=load_workbook(out);s=w['Payment Input']
        assert s.cell(activity_rows(s)['A100'],5).value==expected
        w.close()


def test_old_layout_and_nonsequential_period_ids_preserved(tmp_path):
    source=requirement_source(tmp_path/'source.xlsx')
    w=load_workbook(source);s=w.create_sheet('Payment Input')
    for c,v in enumerate(['Type','WBS','Activity ID','Activity Name','P1','P09'],1):s.cell(6,c,v)
    s.append([]);s['A8']='ACT';s['C8']='A100';s['E8']=0;s['F8']=.37
    _protect_payment_input(s)
    assert not s['E8'].protection.locked and not s['F8'].protection.locked
    w.save(source);w.close()
    before=values(source)
    PaymentService().prepare_embedded_payment_input(source,source,2)
    after=values(source)
    assert list(after)==['P1','P09']
    assert after['P1']['A100']==0 and after['P09']['A100']==.37
    w=load_workbook(source);s=w['Payment Input'];assert s['F6'].value=='P1';w.close()


@pytest.mark.parametrize('failure',['invalid_input','save','replace','remove_period','duplicate_id'])
def test_failed_prepare_preserves_source_and_existing_output(tmp_path,monkeypatch,failure):
    source=requirement_source(tmp_path/'source.xlsx')
    PaymentService().prepare_embedded_payment_input(source,source,3)
    edit_requirements(source)
    if failure in ('invalid_input','duplicate_id'):
        w=load_workbook(source)
        if failure=='invalid_input':w['Payment Input'].cell(activity_rows(w['Payment Input'])['A100'],6,'bad')
        else:w['main']['E9']='A100'
        w.save(source);w.close()
    original=source.read_bytes();out=tmp_path/'existing.xlsx';out.write_bytes(b'KEEP')
    def fail(*a,**kw):raise OSError('simulated failure')
    if failure=='save':monkeypatch.setattr(Workbook,'save',fail)
    if failure=='replace':monkeypatch.setattr('progress_studio.services.payment_service.os.replace',fail)
    with pytest.raises((ValueError,OSError)):
        PaymentService().prepare_embedded_payment_input(source,out,1 if failure=='remove_period' else 3)
    assert source.read_bytes()==original and out.read_bytes()==b'KEEP'
    assert not list(tmp_path.glob('.*.payment-input.*'))


def recalculate(source,directory):
    office=shutil.which('soffice') or shutil.which('libreoffice')
    if not office:pytest.skip('LibreOffice not installed; native Excel remains an acceptance gate')
    directory.mkdir()
    # Exercise explicit recalculation on a disposable copy. Conversion alone
    # honors Manual mode and is not an F9 equivalent.
    incoming = directory/'incoming'
    incoming.mkdir()
    copy_path = incoming/source.name
    wb = load_workbook(source)
    wb.calculation.calcMode = 'auto'
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(copy_path);wb.close()
    # File-backed output avoids inherited pipe handles keeping communicate()
    # alive after an office launcher has finished in headless environments.
    log = directory/'office.log'
    with log.open('w') as stream:
        result=subprocess.run([office,'-env:UserInstallation='+ (directory/'profile').as_uri(),
            '--headless','--convert-to','xlsx','--outdir',str(directory),str(copy_path)],
            stdout=stream,stderr=subprocess.STDOUT,text=True,timeout=180)
    output=directory/source.name
    assert result.returncode==0 and output.exists(),log.read_text()
    return output


def test_calculated_wbs_values_and_live_edits(tmp_path):
    source=requirement_source(tmp_path/'source.xlsx');out=tmp_path/'proof.xlsx'
    PaymentService().prepare_embedded_payment_input(source,out,3);edit_requirements(out)
    calc=recalculate(out,tmp_path/'calc')
    w=load_workbook(calc,data_only=True);s=w['Payment Input']
    wbs={s.cell(r,2).value:r for r in range(8,s.max_row+1) if s.cell(r,1).value=='WBS'}
    r=wbs['1']
    assert s.cell(r,5).value==400
    assert s.cell(r,6).value==pytest.approx(.4)
    assert s.cell(r,7).value==pytest.approx(.15) and s.cell(r,8).value==pytest.approx(.15)
    assert s.cell(wbs['1.2'],6).value==pytest.approx(.2)
    assert s.cell(wbs['2'],6).value==.5 and s.cell(wbs['2'],8).value==.5
    assert s.cell(wbs['3'],6).value=='Check Amount'
    assert s.cell(wbs['4'],6).value=='No weight'
    assert s.cell(wbs['5'],6).value in (None,'')
    w.close()
    w=load_workbook(out);s=w['Payment Input'];rows=activity_rows(s)
    s.cell(rows['A200'],7,'50%');s.cell(rows['A100'],6).value=None
    w.save(out);w.close()
    calc=recalculate(out,tmp_path/'edited')
    w=load_workbook(calc,data_only=True);s=w['Payment Input']
    assert s.cell(r,6).value==pytest.approx(.15)
    assert s.cell(r,7).value==pytest.approx(.375)
    assert s.cell(r,8).value==pytest.approx(.375)
    w.close()


def test_migrates_cov_layout_without_changing_sparse_requirements(tmp_path):
    source=requirement_source(tmp_path/'old.xlsx')
    PaymentService().prepare_embedded_payment_input(source,source,3)
    edit_requirements(source)
    w=load_workbook(source);s=w['Payment Input']
    for r in reversed([r for r in range(8,s.max_row+1) if s.cell(r,1).value=='WBS']):
        s.insert_rows(r+1);s.cell(r+1,1,'COV');s.cell(r+1,4,'Amount coverage')
    w.save(source);w.close();before=values(source)
    PaymentService().prepare_embedded_payment_input(source,source,3)
    assert values(source)==before
    w=load_workbook(source);s=w['Payment Input']
    assert all(s.cell(r,1).value!='COV' for r in range(8,s.max_row+1))
    assert s.max_row==20 and s.freeze_panes=='F8'
    assert len(s.conditional_formatting)>0
    w.close()
    PaymentService().prepare_embedded_payment_input(source,source,3)
    assert values(source)==before


def test_calculated_cumulative_edge_cases_and_zero_format(tmp_path):
    from progress_studio.infrastructure.excel.payment_input_summary import WBS_PERCENT_FORMAT
    source=requirement_source(tmp_path/'source.xlsx');out=tmp_path/'cases.xlsx'
    PaymentService().prepare_embedded_payment_input(source,out,3);edit_requirements(out)
    w=load_workbook(out);base=w['Payment Input']
    # Remove unrelated sheets for a quick independent calculation proof.
    for sheet in list(w):
        if sheet!=base:w.remove(sheet)
    cases={
        'before':{'A100':(None,None,.6),'A200':(None,.2,.3)},
        'unknown':{'A200':(None,None,None)},
        'allblank':{k:(None,None,None) for k in activity_rows(base)},
        'decrease':{'A100':(.8,None,.2),'A200':(.2,None,None),'A400':(None,None,None)},
        'tiny_zero':{'A300':(1e-9,0,None)},
        'invalid':{'A200':(.2,'bad',None)},
        'zero_milestone':{'M1':(None,None,None)},
    }
    for name,updates in cases.items():
        s=w.copy_worksheet(base);s.title=name;rows=activity_rows(s)
        for identity,seq in updates.items():
            for c,value in enumerate(seq,6):s.cell(rows[identity],c).value=value
        if name=='decrease':s.cell(rows['A400'],5,0)  # expose the decrease diagnostic
        if name=='allblank':s.cell(rows['A400'],5,0)
    assert base['F8'].number_format==WBS_PERCENT_FORMAT
    assert WBS_PERCENT_FORMAT.split(';')[2]=='' and WBS_PERCENT_FORMAT.split(';')[0]=='0.0%'
    assert base.cell(activity_rows(base)['A100'],7).number_format=='0%'
    assert base.cell(activity_rows(base)['A100'],7).value==0
    w.save(out);w.close()
    calc=recalculate(out,tmp_path/'calculated_cases');w=load_workbook(calc,data_only=True)
    assert [w['before'].cell(9,c).value for c in range(6,9)]==pytest.approx([0,.15,.375])
    assert all(w['unknown'].cell(9,c).value=='Incomplete' for c in range(6,9))
    assert all(w['allblank'].cell(9,c).value in (None,'') for c in range(6,9))
    assert [w['decrease'].cell(9,c).value for c in range(6,9)]==pytest.approx([.35,.35,.2])
    assert 'Review decreasing requirements: A100' in w['decrease']['D4'].value
    assert [w['tiny_zero'].cell(14,c).value for c in range(6,9)]==pytest.approx([1e-9,0,0],abs=1e-14)
    assert all(w['invalid'].cell(9,c).value=='Check input' for c in range(6,9))
    assert w['zero_milestone']['F9'].value==pytest.approx(.4)
    w.close()
