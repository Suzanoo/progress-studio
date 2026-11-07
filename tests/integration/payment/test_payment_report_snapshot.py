from pathlib import Path
from io import BytesIO
from zipfile import ZipFile,ZIP_DEFLATED
from xml.etree import ElementTree as ET
from unittest.mock import patch
import pytest
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from progress_studio.infrastructure.excel.payment_report_snapshot import inspect_snapshots,append_payment_snapshots
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from progress_studio.services.boq_payment_export import BOQPaymentExportService,digest
from tests.unit.payment.test_boq_payment_projection import run

NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def patch_caches(path, missing=False):
    """Test fixture cached values only; never evaluates production formulas."""
    out=BytesIO()
    with ZipFile(path) as src,ZipFile(out,'w',ZIP_DEFLATED) as dst:
        for info in src.infolist():
            b=src.read(info.filename)
            if info.filename.startswith('xl/worksheets/') and info.filename.endswith('.xml'):
                root=ET.fromstring(b)
                for c in root.findall('.//s:sheetData/s:row/s:c',NS):
                    if c.find('s:f',NS) is None:continue
                    v=c.find('s:v',NS)
                    if v is None:v=ET.SubElement(c,'{'+NS['s']+'}v')
                    if missing:v.text=None;c.attrib.pop('t',None)
                    elif c.attrib['r']=='D2':v.text='0';c.attrib.pop('t',None)
                    elif c.attrib['r']=='D3':v.text='-2.5';c.attrib.pop('t',None)
                    else:c.set('t','str');v.text=None
                b=ET.tostring(root)
            dst.writestr(info,b)
    path.write_bytes(out.getvalue())


def source_fixture(path):
    w=Workbook();w.remove(w.active)
    for name in ('Payment Input','Payment'):
        s=w.create_sheet(name);s['A1']='ข้อมูลการชำระเงิน';s['A2']='ACT';s['B2']='A100';s['C2']=0
        s['D2']='=0';s['D3']='=-2.5';s['D4']='=""';s['E2']='=literal text';s['E2'].data_type='s'
        s['B3']='รายละเอียดภาษาไทย';s['B3'].font=Font(bold=True,color='123456');s['C2'].protection=Protection(locked=False)
        s.row_dimensions[2].outlineLevel=1;s.row_dimensions[2].hidden=True;s.column_dimensions['B'].width=32;s.column_dimensions['G'].hidden=True
        s.freeze_panes='C2';s.print_area='A1:G8';s.print_title_rows='1:1';s.merge_cells('A6:C6')
        s['F2'].hyperlink='https://example.com';s['F2']='source link'
        dv=DataValidation(type='list',formula1='"a,b"');s.add_data_validation(dv);dv.add(s['F3'])
        s.conditional_formatting.add('D2:D4',FormulaRule(formula=['ISNUMBER(D2)'],fill=PatternFill('solid',fgColor='ABCDEF')))
    w.save(path);w.close();patch_caches(path);return path


def test_saved_values_styles_controls_and_unicode(tmp_path):
    path=source_fixture(tmp_path/'ข้อมูล payment.xlsx');before=path.read_bytes();target=render(run())
    infos=append_payment_snapshots(target,path,digest(path));assert [i.formulas for i in infos]==[3,3]
    out=tmp_path/'snapshot.xlsx';target.save(out);target.close();BOQPaymentExportService.validate(out)
    w=load_workbook(out)
    assert w.sheetnames==['Percentage','Amount','Trace','Payment Input','Payment','_BP_Data']
    for name in ('Payment Input','Payment'):
        s=w[name];assert s['D2'].value==0 and s['D3'].value==-2.5 and s['D4'].value is None
        assert s['C2'].value==0 and s['E2'].value=='=literal text' and s['E2'].data_type=='s'
        assert s['B3'].value=='รายละเอียดภาษาไทย' and s['B3'].font.color.rgb=='00123456'
        assert s['C2'].protection.locked and s.protection.sheet
        assert not s['F2'].hyperlink and len(s.data_validations.dataValidation)==0
        assert s.freeze_panes=='C2' and s.row_dimensions[2].outlineLevel==1 and s.column_dimensions['G'].hidden
        assert not s.row_dimensions[2].hidden
        assert 'A6:C6' in s.merged_cells and s['A1'].value.startswith('SNAPSHOT')
        assert digest(path) in s['A1'].comment.text
        assert not any(c.data_type=='f' for row in s for c in row)
    assert path.read_bytes()==before;w.close()


def test_missing_formula_cache_is_not_empty_string(tmp_path):
    path=source_fixture(tmp_path/'input.xlsx');assert inspect_snapshots(path)[0].formulas==3
    patch_caches(path,missing=True)
    with pytest.raises(ValueError,match='Missing saved formula value'):inspect_snapshots(path)


def test_missing_sheet_and_cancellation(tmp_path):
    path=source_fixture(tmp_path/'input.xlsx')
    with pytest.raises(InterruptedError):inspect_snapshots(path,lambda:True)
    w=load_workbook(path);w.remove(w['Payment']);w.save(path);w.close()
    # Missing cache is also rejected; a literal-only missing-sheet fixture follows.
    w=Workbook();w.active.title='Payment Input';w.save(path);w.close()
    with pytest.raises(ValueError,match='worksheet missing: Payment'):inspect_snapshots(path)


def test_foreign_conditional_formatting_fails_closed(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx');w=load_workbook(path)
    w['Payment'].conditional_formatting.add('C2',FormulaRule(formula=["'main'!A1=1"]))
    w.save(path);w.close();patch_caches(path)
    target=render(run())
    with pytest.raises(ValueError,match='formatting dependency'):append_payment_snapshots(target,path,digest(path))
    target.close()


def test_inflated_dimensions_do_not_create_cells(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx');out=BytesIO()
    with ZipFile(path) as src,ZipFile(out,'w',ZIP_DEFLATED) as dst:
        for info in src.infolist():
            b=src.read(info.filename)
            if info.filename=='xl/worksheets/sheet1.xml':
                root=ET.fromstring(b);root.find('s:dimension',NS).set('ref','A1:XFD1048576');b=ET.tostring(root)
            dst.writestr(info,b)
    path.write_bytes(out.getvalue());target=render(run());append_payment_snapshots(target,path,digest(path))
    assert len(target['Payment Input']._cells)<100;target.close()


def test_stale_numeric_cache_preserved_and_not_recomputed(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx')
    # Formula is deliberately inconsistent with its saved numeric value.
    out=BytesIO()
    with ZipFile(path) as src,ZipFile(out,'w',ZIP_DEFLATED) as dst:
        for info in src.infolist():
            b=src.read(info.filename)
            if info.filename=='xl/worksheets/sheet1.xml':b=b.replace(b'>0</ns0:f>',b'>999</ns0:f>')
            dst.writestr(info,b)
    path.write_bytes(out.getvalue());target=render(run());append_payment_snapshots(target,path,digest(path))
    assert target['Payment Input']['D2'].value==0
    assert 'freshness unverified' in target['Payment Input']['A1'].value
    target.close()


def test_source_errors_remain_errors_and_are_counted(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx');w=load_workbook(path);w['Payment']['G5']='#DIV/0!';w.save(path);w.close();patch_caches(path)
    assert inspect_snapshots(path)[1].errors==1
    target=render(run());append_payment_snapshots(target,path,digest(path));assert target['Payment']['G5'].data_type=='e';target.close()


def test_static_chart_series_have_no_source_references(tmp_path):
    from openpyxl.chart import BarChart,Reference
    from openpyxl.chart.data_source import NumData,NumVal
    path=source_fixture(tmp_path/'source.xlsx');w=load_workbook(path);s=w['Payment']
    s.unmerge_cells('A6:C6');s['C5']=3;s['C6']=7;chart=BarChart();chart.add_data(Reference(s,min_col=3,min_row=5,max_row=6))
    chart.series[0].val.numRef.numCache=NumData(ptCount=2,pt=[NumVal(idx=0,v=3),NumVal(idx=1,v=7)])
    s.add_chart(chart,'I2');w.save(path);w.close();patch_caches(path)
    target=render(run());append_payment_snapshots(target,path,digest(path));out=tmp_path/'out.xlsx';target.save(out);target.close()
    BOQPaymentExportService.validate(out)
    with ZipFile(out) as z:
        chart=ET.fromstring(z.read('xl/charts/chart1.xml'))
        assert not any(e.tag.endswith('}f') for e in chart.iter())
        assert any(e.tag.endswith('}numLit') for e in chart.iter())


def test_uncached_chart_rejected(tmp_path):
    from openpyxl.chart import BarChart,Reference
    path=source_fixture(tmp_path/'source.xlsx');w=load_workbook(path);s=w['Payment'];s['C5']=3
    chart=BarChart();chart.add_data(Reference(s,min_col=3,min_row=5,max_row=5));s.add_chart(chart,'I2')
    w.save(path);w.close();patch_caches(path);target=render(run())
    with pytest.raises(ValueError,match='saved chart'):append_payment_snapshots(target,path,digest(path))
    target.close()


def test_snapshot_failure_never_publishes(tmp_path):
    from progress_studio.services.boq_payment_export import ExportReview
    source=source_fixture(tmp_path/'source.xlsx');session=tmp_path/'session';session.write_bytes(b'session')
    review=ExportReview(run(),((session,digest(session)),(source,digest(source))),False)
    before=source.read_bytes();out=tmp_path/'out.xlsx'
    for error in (InterruptedError('cancel in snapshot'),ValueError('unsupported snapshot'),PermissionError('write denied')):
        with patch('progress_studio.services.boq_payment_export.append_payment_snapshots',side_effect=error):
            with pytest.raises(type(error)):BOQPaymentExportService().export(review,out)
        assert not out.exists() and not list(tmp_path.glob('.bp2-*')) and source.read_bytes()==before


def test_review_reports_cache_limits_without_changing_projection(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx');session=tmp_path/'session';session.write_bytes(b'session');projection=run()
    with patch('progress_studio.services.boq_payment_export.BOQPaymentProjectionService.calculate',return_value=projection):
        review=BOQPaymentExportService().review(session,path)
    assert review.projection.items==projection.items
    assert len([d for d in review.projection.diagnostics if d.code=='PaymentSnapshot'])==2
    assert all('freshness is unverified' in d.detail for d in review.projection.diagnostics if d.code=='PaymentSnapshot')


def test_missing_cache_fails_review_before_projection(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx');patch_caches(path,missing=True);session=tmp_path/'session';session.write_bytes(b'session')
    with patch('progress_studio.services.boq_payment_export.BOQPaymentProjectionService.calculate') as calculate:
        with pytest.raises(ValueError,match='Missing saved'):BOQPaymentExportService().review(session,path)
        calculate.assert_not_called()



def test_scrolled_origin_is_not_mistaken_for_frozen_extent(tmp_path):
    path=source_fixture(tmp_path/'source.xlsx');w=load_workbook(path)
    w['Payment Input'].sheet_view.pane.topLeftCell='Z233'
    w['Payment']['B4']='Activity ID';w['Payment'].freeze_panes='AE659'
    w.save(path);w.close();patch_caches(path)
    target=render(run());append_payment_snapshots(target,path,digest(path))
    pane=target['Payment Input'].sheet_view.pane
    assert pane.xSplit==2 and pane.ySplit==1 and pane.topLeftCell=='Z233'
    pane=target['Payment'].sheet_view.pane
    assert pane.xSplit==30 and pane.ySplit==4 and pane.topLeftCell=='AE5'
    target.close()
