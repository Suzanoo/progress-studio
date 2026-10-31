"""Reject non-finite Payment Input at the shared snapshot/live boundary."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
import pytest
from openpyxl import load_workbook
from progress_studio.infrastructure.excel.payment_input_reader import PaymentInputSparseReader
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError, MAIN_NS
from progress_studio.services.rebuild_service import WorkbookRebuildEngine
from tests.integration.payment.test_live_payment_integration import _fixture

def _nonfinite(path, value, numeric):
    _fixture(path)
    wb=load_workbook(path);wb['Payment Input']['E8']=value;wb.save(path);wb.close()
    if numeric:
        # openpyxl normalizes Python NaN to an empty cell. Exercise raw numeric XML.
        with ZipFile(path) as archive:parts={n:archive.read(n) for n in archive.namelist()}
        for name,raw in parts.items():
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                root=ET.fromstring(raw)
                if b'Payment Input' not in raw and b'Activity ID' not in raw:
                    continue
                cell=root.find(f".//{{{MAIN_NS}}}c[@r='E8']")
                if cell is not None:
                    cell.clear();cell.attrib.update(r='E8',t='n')
                    ET.SubElement(cell,f'{{{MAIN_NS}}}v').text=value
                    parts[name]=ET.tostring(root)
        with ZipFile(path,'w',ZIP_DEFLATED) as archive:
            for name,raw in parts.items():archive.writestr(name,raw)
    return path

@pytest.mark.parametrize('value', ['NaN','nan',' NaN ','NaN%','Infinity','-Infinity','inf','-inf','1e309'])
def test_text_nonfinite_rejected_with_cell(tmp_path,value):
    source=_nonfinite(tmp_path/'input.xlsx',value,False)
    with pytest.raises(PaymentWorkbookError,match='E8.*finite percentage'):
        PaymentInputSparseReader().read(source)

@pytest.mark.parametrize('value',['NaN','Infinity','-Infinity'])
@pytest.mark.parametrize('numeric',[False,True])
@pytest.mark.parametrize('live',[False,True])
def test_both_rebuilds_reject_without_overwriting(tmp_path,value,numeric,live):
    source=_nonfinite(tmp_path/'input.xlsx',value,numeric)
    before=source.read_bytes()
    output=tmp_path/'existing.xlsx';output.write_bytes(b'KEEP_EXISTING')
    engine=WorkbookRebuildEngine()
    method=engine.rebuild_live_payment if live else engine.rebuild_payment
    with pytest.raises(PaymentWorkbookError,match='E8.*finite percentage'):
        method(source,output)
    assert source.read_bytes()==before
    assert output.read_bytes()==b'KEEP_EXISTING'
    assert not list(tmp_path.glob('.existing.*'))

@pytest.mark.parametrize('value,expected',[(None,None),(0,0),(1,1),(.5,.5),('50%',.5)])
def test_blank_zero_and_valid_finite_values_unchanged(tmp_path,value,expected):
    source=_nonfinite(tmp_path/'input.xlsx',value,False)
    data=PaymentInputSparseReader().read(source)
    requirements=[r for period in data.periods for r in period.requirements]
    assert [r.required_fraction for r in requirements] == ([] if expected is None else [expected])
