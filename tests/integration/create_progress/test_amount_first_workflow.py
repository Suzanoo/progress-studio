"""Amount-first lifecycle regression; no Desktop Excel acceptance implied."""
from pathlib import Path
import xml.etree.ElementTree as ET
from unittest.mock import Mock
import pytest
from openpyxl import load_workbook

from progress_studio.app.desktop import DesktopRunner, DesktopRunOptions
from progress_studio.app.application import ProgressStudioApplication
from progress_studio.presentation.cli import CommandLineInterface
from progress_studio.config import SETTINGS
from progress_studio.infrastructure.excel.main_dataset_workbook_adapter import main_dataset_from_workbook
from progress_studio.infrastructure.excel.ev_monetary_inputs import SHEET, ACTIVITY, BOQ, prepare_inputs
from progress_studio.infrastructure.excel.weight_basis import (
    amount_initialization, set_amount_initialization, set_creation_basis, creation_basis,
)
from progress_studio.services.earned_value_rebuild_service import EarnedValueRebuildService, EarnedValueRebuildError
from progress_studio.services.rebuild_service import WorkbookRebuildEngine
from progress_studio.services.payment_service import PaymentService
from tests.fixtures.amount_xml import amount_xml

MANUAL = "manual_placeholder"

def _create(tmp_path, monkeypatch, dialect, *, remove_fields=True):
    xml = amount_xml(tmp_path, dialect, ('100','900','0','200'))
    if remove_fields:
        tree = ET.parse(xml)
        for parent in tree.getroot().iter():
            for child in list(parent):
                tag = child.tag.split('}')[-1]
                if tag in ('UDF', 'UDFType'):
                    parent.remove(child)
                elif tag == 'ExtendedAttribute':
                    identity = next((n.text for n in child if n.tag.split('}')[-1] == 'FieldID'), None)
                    if identity != '188743731':
                        parent.remove(child)
        tree.write(xml, encoding='utf-8', xml_declaration=True)
    monkeypatch.setattr('progress_studio.pipeline.import_step.desktop_path', lambda: tmp_path/'desktop')
    return DesktopRunner().run(DesktopRunOptions(
        xml, '5', weight_basis='amount', amount_initialization=MANUAL,
        distribution_method='flat')).output_workbook

def _values(wb, names):
    return {name: {c.coordinate:c.value for row in wb[name] for c in row if c.value is not None}
            for name in names}

@pytest.mark.parametrize('dialect', ['p6','msp'])
@pytest.mark.parametrize('remove_fields', [False,True])
def test_manual_create_blank_bac_and_truthful_origin(tmp_path, monkeypatch, dialect, remove_fields):
    path = _create(tmp_path, monkeypatch, dialect, remove_fields=remove_fields)
    wb = load_workbook(path)
    assert [a.amount for a in main_dataset_from_workbook(wb).activities] == [1,1,1,0]
    assert creation_basis(wb) == 'amount'
    assert amount_initialization(wb) == MANUAL
    assert [wb[SHEET].cell(r,3).value for r in range(4,8)] == [None]*4
    assert wb[SHEET]['D7'].value is True
    assert not {'Payment', 'Earned Value', 'EV Table', 'EV_Data'} & set(wb.sheetnames)
    assert 'Progress Rebuild' in wb['README']['B4'].value
    assert 'monetary XML values' not in wb['README']['B4'].value
    wb.close()
    with pytest.raises(EarnedValueRebuildError, match='monetary Amount'):
        EarnedValueRebuildService().analyze(path, monetary_source=ACTIVITY)

@pytest.mark.parametrize('dialect', ['p6','msp'])
def test_payment_ev_and_progress_preserve_independent_inputs(tmp_path, monkeypatch, dialect):
    source = _create(tmp_path, monkeypatch, dialect)
    payment = PaymentService()
    prepared = tmp_path/'input.xlsx'
    payment.prepare_embedded_payment_input(source, prepared, periods=1)
    wb = load_workbook(prepared)
    ds = main_dataset_from_workbook(wb)
    for i,activity in enumerate(ds.activities[:3]):
        for j,period in enumerate(ds.periods):
            wb['main'].cell(activity.row_number, period.column, float(j == i % len(ds.periods)))
            wb['main'].cell(activity.row_number+1, period.column, .25 if j==0 else 0)
    pin = wb['Payment Input']
    headers = {c.value:c.column for c in pin[6]}
    for r in range(8,pin.max_row+1):
        pin.cell(r,headers['P01']).value = .5 if pin.cell(r,headers['Activity ID']).value == ds.activities[0].activity_id else None
    wb.create_sheet('User Notes')['A1'] = 'keep'
    names = ['main','Payment Input',SHEET,'User Notes']
    before = _values(wb,names)
    wb.save(prepared);wb.close()
    engine = WorkbookRebuildEngine()
    for mode,method in [('snapshot',engine.rebuild_payment),('live',engine.rebuild_live_payment)]:
        out = tmp_path/f'payment-{mode}.xlsx'
        result = method(prepared,out)
        assert result.rendered_points == 1
        wb = load_workbook(out)
        assert 'BOQ Activity Mapping' not in wb.sheetnames
        assert _values(wb,names) == before
        assert all(wb[SHEET].cell(r,3).value is None for r in range(4,8))
        wb.close()
    wb = load_workbook(prepared)
    for row,value in enumerate((100,900,0,200),4):wb[SHEET].cell(row,3,value)
    wb.save(prepared);wb.close()
    ev = EarnedValueRebuildService()
    assert ev.analyze(prepared,monetary_source=ACTIVITY).project_bac == 1200
    with pytest.raises(EarnedValueRebuildError):
        ev.analyze(prepared,monetary_source=BOQ)
    evout = tmp_path/'ev.xlsx'
    ev.generate(prepared,evout,monetary_source=ACTIVITY)
    wb = load_workbook(evout)
    col = ds.header_column('amount'); row = ds.activities[0].row_number
    wb['main'].cell(row,col,999)
    assert wb['main_monthly'].cell(row,col).value == 1
    before = _values(wb,names)
    wb.save(evout);wb.close()
    for mode,method in [('snapshot',engine.rebuild_progress),('live',engine.rebuild_live_progress)]:
        out = tmp_path/f'progress-{mode}.xlsx';method(evout,out)
        wb = load_workbook(out)
        assert _values(wb,names) == before
        assert amount_initialization(wb) == MANUAL
        assert wb['main_monthly'].cell(row,col).value == 999
        assert 'monetary XML values' not in wb['README']['B4'].value
        if mode == 'live':
            assert wb['Dashboard']['J6'].value == 'Total weight'
            assert wb['Dashboard']['K6'].value == 1001
        wb.close()
        assert ev.analyze(out,monetary_source=ACTIVITY).project_bac == 1200

@pytest.mark.parametrize('origin,basis,expected', [
    (MANUAL,'amount',None),('xml_field','amount',37),(None,'amount',37),
    ('unsupported','amount',None),(None,'equal',None),(None,'duration',None),(None,None,None),
])
def test_legacy_preparation_respects_origin_after_main_edits(tmp_path, origin, basis, expected):
    from tests.integration.create_progress.test_ev_monetary_sources import source_book
    source = source_book(tmp_path/'source.xlsx')
    wb=load_workbook(source);del wb[SHEET];wb['main']['O9']=37
    if basis:set_creation_basis(wb,basis)
    if origin:set_amount_initialization(wb,origin)
    wb.save(source);wb.close()
    with pytest.raises(EarnedValueRebuildError,match='Missing'):
        EarnedValueRebuildService().analyze(source,monetary_source=ACTIVITY)
    out=tmp_path/'prepared.xlsx';prepare_inputs(source,out)
    wb=load_workbook(out)
    assert wb[SHEET]['C4'].value == expected
    wb.close()

@pytest.mark.parametrize('basis,initialization,field', [
    ('amount','xml_field',None),('equal',MANUAL,None),
    ('amount',MANUAL,'p6:udf:91'),('amount','unknown',None),
])
def test_invalid_creation_intent_never_falls_back(tmp_path,basis,initialization,field):
    xml=amount_xml(tmp_path,'p6')
    runner=DesktopRunner(pipeline_factory=Mock(side_effect=AssertionError('Pipeline must not run')))
    with pytest.raises(ValueError):
        runner.run(DesktopRunOptions(xml,'5',weight_basis=basis,amount_field=field,
                                     amount_initialization=initialization))

def test_cli_manual_option_reaches_pipeline(tmp_path):
    xml=amount_xml(tmp_path,'msp')
    pipeline=Mock()
    pipeline.run.return_value.output_workbook=None
    app=ProgressStudioApplication(pipeline,CommandLineInterface(SETTINGS))
    assert app.run(['--input',str(xml),'--cutoff-day','5','--weight-basis','amount',
                    '--amount-initialization',MANUAL]) == 0
    context=pipeline.run.call_args.args[0]
    assert context.amount_initialization == MANUAL and context.amount_field is None
