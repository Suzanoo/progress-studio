from dataclasses import replace
from decimal import Decimal as D
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile
import pytest
from openpyxl import load_workbook
from progress_studio.domain.boq_manual_requirements import ManualInput, compose
from progress_studio.domain.boq_payment_report import known_projected_totals, report_rows, amount_completion_states
from progress_studio.domain.boq_payment_projection import project
from progress_studio.domain.mapping_models import AllocationRecord
from progress_studio.infrastructure.excel.boq_requirement_input import read_input, render_input, SHEET
from progress_studio.infrastructure.excel.boq_input_package import merge_input
from progress_studio.services.boq_requirement_prepare import BOQRequirementPrepareService
from progress_studio.services.boq_payment_export import BOQPaymentExportService, ExportReview, digest
from tests.unit.payment.test_boq_payment_projection import run, item, payment
from tests.integration.payment.test_boq_payment_export import reviewed


def manual(values, amount=D(100)):
    p=run(shares=(),amount=amount)
    return compose(p,(ManualInput('b',tuple(None if v is None else D(str(v)) for v in values),9),))


def test_signed_period_unknown_prefix_and_zero():
    p=manual((.5,0,-.1));row=p.items[0]
    assert [v.cumulative_fraction for v in row.periods]==[D('.5'),D('.5'),D('.4')]
    assert known_projected_totals(p).period==(D(50),D(0),D(-10))
    assert known_projected_totals(p).cumulative==(D(50),D(50),D(40))
    p=manual((.5,None,.1))
    assert known_projected_totals(p).cumulative==(D(50),None,None)
    assert known_projected_totals(p).period==(D(50),None,D(10))
    assert report_rows(p)[0].values[1][-1]==D('.1')
    assert report_rows(p)[0].values[0][-1] is None
    assert p.items[0].periods[2].state=='Incomplete history'
    assert 'Incomplete' in known_projected_totals(p).status
    assert all(not v.contributions and v.mapping_coverage==0 for v in p.items[0].periods)


@pytest.mark.parametrize('values',[(-.1,0,0),(.9,.2,0),(1.01,0,0),(0,float('inf'),0),(0,float('nan'),0)])
def test_invalid_numeric_or_prefix_rejected(values):
    with pytest.raises(ValueError):manual(values)


@pytest.mark.parametrize('amount',[None,D(-1),D(0)])
def test_amount_does_not_manufacture_requirement(amount):
    p=manual((None,0,.1),amount)
    assert p.items[0].periods[0].cumulative_fraction is None
    assert p.items[0].periods[2].incremental_amount==(D(0) if amount==0 else None)
    assert not any(any(v) for v in amount_completion_states(p,report_rows(p)))


def test_disjoint_origins_and_no_double_counting():
    a=item('a');b=replace(item('b',D(200)),wbs='1.2')
    p=project((a,b),(AllocationRecord('a','A',100),),payment(),('A','B'))
    result=compose(p,(ManualInput('b',(D('.1'),D(0),D('.2')),9),))
    assert result.items[0] is p.items[0] and result.mapped_projection is p
    assert known_projected_totals(result).cumulative==(D(40),D(40),D(140))
    assert known_projected_totals(result).period==(D(40),D(0),D(100))
    assert report_rows(result)[0].amount==300
    assert compose(p,None) is p
    with pytest.raises(ValueError,match='conflict'):compose(p,(ManualInput('a',(D(0),)*3,9),))
    with pytest.raises(ValueError,match='Duplicate'):compose(p,(ManualInput('b',(D(0),)*3,9),)*2)


def test_exact_completion_and_decrease():
    p=manual((1,-.1,.1));assert amount_completion_states(p,report_rows(p))==((True,False,True),)*2


def prepared(tmp_path, p=None):
    p=p or run(shares=());r=reviewed(tmp_path,p);path=tmp_path/'prepared.xlsx'
    BOQRequirementPrepareService().export(r,path,acknowledge_revision=True)
    return p,r,path


def test_prepare_preserves_original_package_and_editable_input(tmp_path):
    p,r,out=prepared(tmp_path)
    source=r.sources[1][0]
    allowed={'xl/styles.xml','xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml'}
    with ZipFile(source) as a,ZipFile(out) as b:
        assert all(a.read(n)==b.read(n) for n in a.namelist() if n not in allowed)
    records=read_input(out,p);assert records[0].values==(None,)*3
    w=load_workbook(out);s=w[SHEET]
    assert s.protection.sheet and not s['I9'].protection.locked and s['G9'].protection.locked
    assert s.freeze_panes=='I9' and s.sheet_state=='visible'
    s['I9']=.5;s['J9']=0;s['K9']=-.1;w.save(out);w.close()
    records=read_input(out,p);assert records[0].values==(D('.5'),D(0),D('-.1'))
    second=tmp_path/'second.xlsx';r2=replace(r,sources=(r.sources[0],(out,digest(out))))
    BOQRequirementPrepareService().export(r2,second,acknowledge_revision=True)
    assert read_input(second,p)==records
    assert all(digest(path)==h for path,h in r.sources)


@pytest.mark.parametrize('cell,value',[('I9','=0.1'),('I9','10%'),('I9',True),('I7','P99'),('A9','bad'),('C9','bad'),('L9',.5),('B2','{}')])
def test_corrupt_schema_identity_or_input_rejected(tmp_path,cell,value):
    p,r,out=prepared(tmp_path);w=load_workbook(out);w[SHEET][cell]=value;w.save(out);w.close()
    with pytest.raises(ValueError):read_input(out,p)


def test_new_mapping_conflict_and_source_revision(tmp_path):
    p,r,out=prepared(tmp_path)
    with pytest.raises(ValueError):read_input(out,run(),reconcile=True)
    changed=replace(p,items=(replace(p.items[0],item=replace(p.items[0].item,amount=D(120))),))
    with pytest.raises(ValueError,match='revision'):read_input(out,changed)
    assert read_input(out,changed,reconcile=True)[0].values==(None,)*3


@pytest.mark.parametrize('failure',['existing','alias','changed','ack','cancel','save','merge','publish'])
def test_prepare_failure_safety(tmp_path,failure):
    r=reviewed(tmp_path,run(shares=()));out=tmp_path/'out.xlsx';service=BOQRequirementPrepareService()
    if failure=='existing':out.write_bytes(b'keep')
    if failure=='alias':out=r.sources[1][0]
    if failure=='changed':r.sources[0][0].write_bytes(b'changed')
    names={'save':'openpyxl.workbook.workbook.Workbook.save','merge':'progress_studio.services.boq_requirement_prepare.merge_input','publish':'progress_studio.services.boq_requirement_prepare.os.link'}
    if failure in names:
        with patch(names[failure],side_effect=PermissionError('simulated failure')):
            with pytest.raises(PermissionError):service.export(r,out,acknowledge_revision=True)
    else:
        with pytest.raises((ValueError,FileExistsError,InterruptedError)):
            service.export(r,out,acknowledge_revision=failure!='ack',cancelled=lambda:failure=='cancel')
    if failure=='existing':assert out.read_bytes()==b'keep'
    elif failure!='alias':assert not out.exists()
    assert not list(tmp_path.glob('.bp4-*'))


def test_manual_export_trace_and_immutable_snapshots(tmp_path):
    p=manual((.5,None,-.1));r=reviewed(tmp_path,p);out=tmp_path/'report.xlsx'
    BOQPaymentExportService().export(r,out,acknowledge_revision=True)
    w=load_workbook(out)
    assert w.sheetnames==['Percentage','Amount','BOQ Mapping','Activity Mapping','Payment Input','Payment','_BP_Data']
    assert 'mapped + manual' in w['Amount']['A5'].value
    assert 'BP_ManualIndex' in w.defined_names
    assert any(c.value=='BOQ Requirement Input!K9' for row in w['_BP_Data'] for c in row)
    assert any(c.value=='Manual period' for row in w['Amount'] for c in row)
    w.close()


@pytest.mark.parametrize('live',[False,True])
def test_actual_payment_rebuild_preserves_manual_input(tmp_path,live):
    from tests.integration.rebuild.test_progress_rebuild_engine import _full_rebuild_fixture
    from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
    from progress_studio.services.rebuild_service import WorkbookRebuildEngine
    source=_full_rebuild_fixture(tmp_path/'source.xlsx')
    w=load_workbook(source);PaymentInputWorkbook().embed(w,periods=3)
    for row in w['Payment Input'].iter_rows(min_row=8):
        if row[0].value=='ACT':row[5].value=.1
    w.save(source);w.close()
    p=run(shares=());sheet=render_input(p,(ManualInput('b',(D('.2'),D(0),None),9),))
    ready=tmp_path/'ready.xlsx';sheet.save(ready);sheet.close();merge_input(source,ready)
    out=tmp_path/'rebuilt.xlsx';service=WorkbookRebuildEngine()
    (service.rebuild_live_payment if live else service.rebuild_payment)(ready,out)
    assert read_input(out,p)[0].values==(D('.2'),D(0),None)
    w=load_workbook(out);assert w[SHEET].sheet_state=='visible' and not w[SHEET]['I9'].protection.locked;w.close()


def test_mapping_save_as_preserves_manual_sheet(tmp_path):
    from tests.integration.mapping.test_mapping_overlay_export import _created_mapping_input
    from progress_studio.services.workbook_export_service import WorkbookExportService
    source,store=_created_mapping_input(tmp_path,True)
    p=run(shares=());w=render_input(p,(ManualInput('b',(D('.2'),D(0),None),9),))
    ready=tmp_path/'ready.xlsx';w.save(ready);w.close();merge_input(source,ready)
    out=tmp_path/'mapped.xlsx';WorkbookExportService().export(ready,out,store)
    assert read_input(out,p)[0].values==(D('.2'),D(0),None)
    w=load_workbook(out);assert w[SHEET].sheet_state=='visible' and not w[SHEET]['I9'].protection.locked;w.close()


def test_period_only_known_status_and_coverage():
    p=manual((None,0,.1))
    assert known_projected_totals(p).period==(None,D(0),D(10))
    assert 'Incomplete' in known_projected_totals(p).status
    diagnostics=[d for d in p.diagnostics if d.code=='ManualCoverage']
    assert len(diagnostics)==3
    assert '1/1 rows with known Period %' in diagnostics[2].detail
    assert '0/1 with known cumulative prefix' in diagnostics[2].detail


def test_excel_shared_strings_survive_mapping_package_copy(tmp_path):
    from zipfile import ZIP_DEFLATED
    from xml.etree import ElementTree as E
    from progress_studio.infrastructure.excel.finance_package import _sheets,S
    from progress_studio.infrastructure.excel.boq_input_package import preserve_input
    p,r,source=prepared(tmp_path)
    with ZipFile(source) as z:package={n:z.read(n) for n in z.namelist()}
    path=_sheets(package)[SHEET][1];root=E.fromstring(package[path]);sst=E.Element('{'+S+'}sst')
    for cell in root.iter('{'+S+'}c'):
        if cell.get('t')=='inlineStr':
            inline=cell.find('{'+S+'}is');cell.remove(inline);inline.tag='{'+S+'}si';sst.append(inline)
            cell.set('t','s');E.SubElement(cell,'{'+S+'}v').text=str(len(sst)-1)
    package[path]=E.tostring(root);package['xl/sharedStrings.xml']=E.tostring(sst)
    with ZipFile(source,'w',ZIP_DEFLATED) as z:
        for n,b in package.items():z.writestr(n,b)
    target=tmp_path/'target.xlsx';target.write_bytes(r.sources[1][0].read_bytes())
    preserve_input(source,target)
    assert read_input(target,p)[0].key=='b'
    w=load_workbook(target);assert w[SHEET]['D1'].value.startswith('BOQ Planned');w.close()
