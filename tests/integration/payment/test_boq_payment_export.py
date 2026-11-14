from dataclasses import replace
from decimal import Decimal as D
from pathlib import Path
from unittest.mock import patch
import os
import pytest
from openpyxl import load_workbook
from progress_studio.services.boq_payment_export import BOQPaymentExportService,ExportReview,digest
from progress_studio.domain.boq_payment_report import report_rows
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from tests.unit.payment.test_boq_payment_projection import run,item,payment
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.mapping_models import AllocationRecord
from tests.integration.payment.test_boq_payment_projection_service import fixture
from tests.integration.payment.test_payment_report_snapshot import source_fixture, patch_caches


def reviewed(tmp_path,projection=None):
    source=tmp_path/'source.bin';source.write_bytes(b'original')
    payment=source_fixture(tmp_path/'payment.xlsx')
    return ExportReview(projection or run(),((source,digest(source)),(payment,digest(payment))),True)


def test_four_matrices_and_schema(tmp_path):
    result=run();r=reviewed(tmp_path,result);out=tmp_path/'report.xlsx'
    BOQPaymentExportService().export(r,out,acknowledge_revision=True)
    w=load_workbook(out);assert w.sheetnames==['Percentage','Amount','BOQ Mapping','Activity Mapping','Payment Input','Payment','_BP_Data']
    rows=report_rows(result);n=len(rows);d=w['_BP_Data']
    for mode in range(4):
        for i,row in enumerate(rows):
            for p,v in enumerate(row.values[mode]):assert d.cell(1+mode*n+i,p+1).value==float(v) if v is not None else d.cell(1+mode*n+i,p+1).value=='Unavailable'
    assert d.sheet_state=='hidden' and d.protection.sheet
    assert not w['Percentage'].protection.sheet and w['Percentage'].freeze_panes=='G8'
    assert w['Percentage']['B3'].value=='Cumulative %';assert w['Amount']['B3'].value=='Cumulative Amount'
    assert len(w['Percentage'].data_validations.dataValidation)==1
    assert 'Payment Trace' not in w.sheetnames
    assert {str(rule.sqref) for rule in w['Percentage'].conditional_formatting} == {'H5', 'I5', 'H8:H9', 'I8:I9'}
    assert {str(rule.sqref) for rule in w['Amount'].conditional_formatting} == {'H8:H9', 'I8:I9'}
    assert w.calculation.calcMode=='manual' and w.calculation.calcOnSave
    assert 'Invalid view' in w['Percentage']['G8'].value
    assert 'ISNUMBER' in next(iter(w['Percentage'].conditional_formatting._cf_rules.values()))[0].formula[0]
    assert not any(c.data_type=='f' for row in d for c in row)
    w.close()


@pytest.mark.parametrize('failure',['ack','existing','source','changed','cancel','save','publish'])
def test_failures_preserve_sources_and_destinations(tmp_path,failure):
    r=reviewed(tmp_path);out=tmp_path/'out.xlsx';before=r.sources[0][0].read_bytes();service=BOQPaymentExportService()
    ack=True;cancel=lambda:False
    if failure=='ack':ack=False
    if failure=='existing':out.write_bytes(b'keep')
    if failure=='source':out=r.sources[0][0]
    if failure=='changed':r.sources[0][0].write_bytes(b'changed')
    if failure=='cancel':cancel=lambda:True
    if failure in ('save','publish'):
        target='openpyxl.workbook.workbook.Workbook.save' if failure=='save' else 'progress_studio.services.boq_payment_export.os.link'
        with patch(target,side_effect=PermissionError('simulated locked/write/publication failure')):
            with pytest.raises(PermissionError):service.export(r,out,acknowledge_revision=ack,cancelled=cancel)
    else:
        with pytest.raises((ValueError,FileExistsError,InterruptedError)):service.export(r,out,acknowledge_revision=ack,cancelled=cancel)
    assert r.sources[0][0].read_bytes()==(b'changed' if failure=='changed' else before)
    if failure=='existing':assert out.read_bytes()==b'keep'
    elif failure!='source':assert not out.exists()
    assert not list(tmp_path.glob('.bp2-*'))


def test_destination_race_and_hardlink(tmp_path):
    r=reviewed(tmp_path);out=tmp_path/'out.xlsx';service=BOQPaymentExportService()
    def race(src,dst):Path(dst).write_bytes(b'other writer');raise FileExistsError('race')
    with patch('progress_studio.services.boq_payment_export.os.link',side_effect=race):
        with pytest.raises(FileExistsError):service.export(r,out,acknowledge_revision=True)
    assert out.read_bytes()==b'other writer'
    link=tmp_path/'alias.xlsx';os.link(r.sources[0][0],link)
    with pytest.raises(ValueError,match='aliases'):service.export(r,link,acknowledge_revision=True)


def test_group_incomplete_fixed_denominator_and_zero():
    a=item('a',D(100));b=replace(item('b',D(200)),wbs='1.2');g=replace(item('g',D(999),'group'),wbs='1')
    p=project((g,a,b),(AllocationRecord('a','A',100),AllocationRecord('b','B',100)),payment(),('A','B'))
    rows=report_rows(p);assert len(rows)==3 and rows[0].amount==300
    assert rows[0].values[0][-1]==D(280)/300
    partial=project((g,a,b),(AllocationRecord('a','A',100),AllocationRecord('b','B',50)),payment(),('A','B'))
    assert report_rows(partial)[0].values[0][-1] is None
    with pytest.raises(ValueError,match='duplicate'):report_rows(replace(p,items=(p.items[0],p.items[1],replace(p.items[2],item=a))))


def test_literal_text_and_unavailable(tmp_path):
    p=run(b=(None,None,None));p=replace(p,items=(replace(p.items[0],item=replace(p.items[0].item,description='=HYPERLINK("bad")')),))
    w=render(p);assert w['Percentage']['B9'].data_type=='s';assert w['_BP_Data']['A1'].value=='Unavailable';w.close()


def test_real_reader_review_and_revision_ack(tmp_path):
    session,current,boq=fixture(tmp_path)
    w=load_workbook(current);w.create_sheet('Payment');w.save(current);w.close();patch_caches(current)
    service=BOQPaymentExportService();r=service.review(session,current,boq)
    assert r.revision_required
    with pytest.raises(ValueError,match='Acknowledge'):service.export(r,tmp_path/'out.xlsx')
    service.export(r,tmp_path/'out.xlsx',acknowledge_revision=True)
    assert all(digest(p)==h for p,h in r.sources)


@pytest.mark.parametrize('amount,expected',[(D(0),'No positive BOQ Amount'),(None,'Invalid Amount'),(D(-1),'Invalid Amount')])
def test_group_zero_invalid(amount,expected):
    assert expected in report_rows(run(amount=amount))[0].status


def test_cancel_during_render_and_before_publication(tmp_path):
    r=reviewed(tmp_path);calls=0
    def cancel():
        nonlocal calls
        calls+=1;return calls>2
    with pytest.raises(InterruptedError):BOQPaymentExportService().export(r,tmp_path/'out.xlsx',acknowledge_revision=True,cancelled=cancel)
    assert not list(tmp_path.glob('.bp2-*')) and not (tmp_path/'out.xlsx').exists()


def test_raw_blank_sentinel_and_all_contributors():
    w=render(run());n=len(report_rows(run()));d=w['_BP_Data'];meta=4*n+2;start=meta+3
    # A and B are retained for every period. B before-first raw cell is blank,
    # its derived requirement is numeric zero. Neither becomes an explicit entry.
    assert d.cell(start+1,3).value=='[blank]' and d.cell(start+1,4).value==0
    assert d.cell(meta,4).value==2
    w.close()

