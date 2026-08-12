from pathlib import Path
from zipfile import ZipFile
from openpyxl import Workbook

from progress_studio.infrastructure.excel.traditional_overlay_workbook import (
    _overlay_chart,
    _responsive_anchor,
)


def test_overlay_module_contract_is_wired():
    text = Path('progress_studio/services/rebuild_service.py').read_text()
    assert 'build_traditional_overlays(wb, dataset)' in text


def test_overlay_uses_all_markers_and_cutoff_helpers():
    text = Path('progress_studio/infrastructure/excel/traditional_overlay_workbook.py').read_text()
    assert 'Weekly Actual Visible' in text
    assert 'Monthly Actual Visible' in text
    assert 'series.marker.symbol = "circle"' in text
    assert 'OVERLAY_MARKER_SIZE = 7' in text
    assert 'DataLabelList(' in text
    assert 'showVal=True' in text
    assert 'numFmt=OVERLAY_LABEL_FORMAT' in text
    assert 'PLAN_LABEL_FILL = "DDEBF7"' in text
    assert 'ACTUAL_LABEL_FILL = "E2F0D9"' in text
    assert 'dLblPos=position' in text
    assert 'spPr=_label_graphical_properties' in text
    assert 'txPr=_label_text_properties' in text
    assert 'chart.x_axis.tickLblPos = "none"' in text
    assert 'Dashboard!$K$5' in text
    assert 'date_col=1' in text and 'plan_col=2' in text and 'actual_col=16' in text
    assert 'date_col=4' in text and 'plan_col=5' in text and 'actual_col=17' in text


def test_overlay_lw1231_responsive_geometry_and_transparency_contract():
    text = Path('progress_studio/infrastructure/excel/traditional_overlay_workbook.py').read_text()
    assert 'OVERLAY_TOP_ROW = 5' in text
    assert 'TwoCellAnchor' in text
    assert 'editAs="twoCell"' in text
    assert '_weekly_project_window(dataset)' in text
    assert '_scurve_plan_row(dataset)' in text
    assert 'scurve_plan_row - 1' in text
    assert 'chart.plot_area.graphicalProperties' in text
    assert 'LineProperties(noFill=True)' in text


def test_overlay_serializes_two_cell_anchor_transparency_and_value_labels(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Dashboard_Data'
    ws.append(['Date', 'Plan', 'Actual'])
    for i in range(1, 5):
        ws.append([i, i / 4, i / 5])

    chart = _overlay_chart(
        data_ws=ws,
        date_col=1,
        plan_col=2,
        actual_col=3,
        first_row=2,
        last_row=5,
    )
    chart.anchor = _responsive_anchor(first_col=4, last_col=7, top_row=5, bottom_row=20)
    ws.add_chart(chart)
    path = tmp_path / 'overlay.xlsx'
    wb.save(path)

    with ZipFile(path) as zf:
        drawing_xml = zf.read('xl/drawings/drawing1.xml').decode('utf-8')
        chart_xml = zf.read('xl/charts/chart1.xml').decode('utf-8')

    assert '<twoCellAnchor editAs="twoCell">' in drawing_xml
    assert '<from><col>3</col>' in drawing_xml
    assert '<to><col>7</col>' in drawing_xml
    # One noFill is the inner plot area and one is the outer chart area.
    assert chart_xml.count('<a:noFill') >= 2
    assert '<dLbls>' in chart_xml
    assert '<showVal val="1"/>' in chart_xml
    assert '<numFmt formatCode="0.0%"' in chart_xml
    # LW-12.3.2 label tags: Plan is above with pale-blue fill; Actual is below
    # with pale-green fill. Text is compact 7 pt and tinted by series.
    assert '<dLblPos val="t"/>' in chart_xml
    assert '<dLblPos val="b"/>' in chart_xml
    assert 'val="DDEBF7"' in chart_xml
    assert 'val="E2F0D9"' in chart_xml
    assert 'val="1F4E79"' in chart_xml
    assert 'val="385723"' in chart_xml
    assert 'sz="700"' in chart_xml


def test_overlay_lw1233_project_bounds_and_cutoff_controls_contract():
    text = Path('progress_studio/infrastructure/excel/traditional_overlay_workbook.py').read_text()
    assert '_project_dates(dataset)' in text
    assert 'p.reporting_date >= start' in text
    assert 'p.reporting_date >= finish' in text
    assert 'first_row=weekly_first' in text
    assert 'last_row=weekly_last' in text
    assert 'first_row=monthly_first' in text
    assert 'last_row=monthly_last' in text
    assert 'label.value = "Cutoff Date"' in text
    assert 'value.value = "=Dashboard!$K$5"' in text
    assert 'Clear the cell to fall back to Dashboard Cutoff' in text
    assert 'list_col="J"' in text
    assert 'list_col="K"' in text
