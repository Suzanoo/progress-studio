from pathlib import Path
import pytest
from openpyxl import load_workbook

from progress_studio.app.desktop import DesktopRunner, DesktopRunOptions
from progress_studio.infrastructure.excel.import_workbook_writer import ImportWorkbookWriter
from progress_studio.infrastructure.excel.main_dataset_workbook_adapter import main_dataset_from_workbook
from progress_studio.infrastructure.excel.weight_basis import creation_basis,creation_field,uses_dummy_weights
from progress_studio.infrastructure.schedule_xml import NormalizedScheduleXmlReader
from progress_studio.services.import_service import ImportService
from progress_studio.services.schedule_service import ScheduleService
from progress_studio.services.rebuild_service import WorkbookRebuildEngine
from tests.fixtures.amount_xml import amount_xml


@pytest.mark.parametrize('source',['p6','msp'])
def test_amount_create_and_both_rebuilds_preserve_values_and_metadata(tmp_path,monkeypatch,source):
    xml=amount_xml(tmp_path,source)
    field=NormalizedScheduleXmlReader().read_with_amount_fields(xml)[2][0]
    monkeypatch.setattr('progress_studio.pipeline.import_step.desktop_path',lambda:tmp_path/'desktop')
    result=DesktopRunner().run(DesktopRunOptions(xml,'5',weight_basis='amount',amount_field=field.identity,distribution_method='flat'))
    outputs=[result.output_workbook]
    for live in (False,True):
        engine=WorkbookRebuildEngine();output=tmp_path/f'rebuild-{live}.xlsx'
        method=engine.rebuild_live_progress if live else engine.rebuild_progress
        method(result.output_workbook,output)
        outputs.append(output)
    for path in outputs:
        wb=load_workbook(path)
        try:
            assert creation_basis(wb)=='amount'
            assert creation_field(wb)==field
            assert not uses_dummy_weights(wb)
            assert 'monetary XML values' in wb['README']['B4'].value
            assert 'BOQ Mapping' not in wb['README']['B4'].value
            assert [r.amount for r in main_dataset_from_workbook(wb).activities]==pytest.approx([10.1234567890123,20.9876543210987,0,0],rel=1e-14)
            assert wb['main'].protection.sheet and wb['Dashboard'].protection.sheet
            assert len(wb['main']._charts)==len(wb['Dashboard']._charts)==1
            assert wb.calculation.calcMode=='manual' and wb.calculation.calcOnSave
        finally:wb.close()


@pytest.mark.parametrize('source',['p6','msp'])
def test_invalid_amount_does_not_create_or_overwrite_output(tmp_path,source):
    xml=amount_xml(tmp_path,source,('-1','10','0','0'))
    service=ImportService(NormalizedScheduleXmlReader(),ScheduleService(),ImportWorkbookWriter())
    field=NormalizedScheduleXmlReader().read_with_amount_fields(xml)[2][0]
    output=tmp_path/'existing.xlsx';output.write_bytes(b'keep me')
    with pytest.raises(ValueError,match='negative'):
        service.import_xml(xml,output,weight_basis='amount',amount_field=field.identity)
    assert output.read_bytes()==b'keep me'
    with pytest.raises(ValueError,match='Select an XML Amount field'):
        service.import_xml(xml,tmp_path/'missing.xlsx',weight_basis='amount')
    assert not (tmp_path/'missing.xlsx').exists()
