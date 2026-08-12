from pathlib import Path
from zipfile import ZipFile
from openpyxl import Workbook

from progress_studio.infrastructure.excel.traditional_overlay_workbook import (
    build_traditional_overlays,
    _overlay_chart,
)


def test_overlay_module_contract_is_wired():
    text = Path('progress_studio/services/rebuild_service.py').read_text()
    assert 'build_traditional_overlays(wb, dataset)' in text


def test_overlay_uses_all_markers_and_cutoff_helpers():
    text = Path('progress_studio/infrastructure/excel/traditional_overlay_workbook.py').read_text()
    assert 'Weekly Actual Visible' in text
    assert 'Monthly Actual Visible' in text
    assert 'series.marker.symbol = "circle"' in text
    assert 'chart.x_axis.tickLblPos = "none"' in text
    assert 'Dashboard!$K$5' in text
    assert 'date_col=1, plan_col=2, actual_col=16' in text
    assert 'date_col=4, plan_col=5, actual_col=17' in text


def test_overlay_lw123_lw124_geometry_and_transparency_contract():
    text = Path('progress_studio/infrastructure/excel/traditional_overlay_workbook.py').read_text()
    assert 'OVERLAY_ANCHOR_ROW = 5' in text
    assert 'chart.plot_area.graphicalProperties' in text
    assert 'LineProperties(noFill=True)' in text
    assert 'oneCellAnchor' in text
    assert 'OVERLAY_HEIGHT_CM' in text
    assert 'OVERLAY_MAX_WIDTH_CM' in text


def test_overlay_serializes_one_cell_anchor_and_two_transparent_layers(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Dashboard_Data'
    ws.append(['Date', 'Plan', 'Actual'])
    for i in range(1, 5):
        ws.append([i, i / 4, i / 5])

    chart = _overlay_chart(
        data_ws=ws, date_col=1, plan_col=2, actual_col=3,
        last_row=5, period_count=4,
    )
    ws.add_chart(chart, 'D5')
    path = tmp_path / 'overlay.xlsx'
    wb.save(path)

    with ZipFile(path) as zf:
        drawing_xml = zf.read('xl/drawings/drawing1.xml').decode('utf-8')
        chart_xml = zf.read('xl/charts/chart1.xml').decode('utf-8')

    assert '<oneCellAnchor>' in drawing_xml
    # One noFill is the inner plot area and one is the outer chart area.
    assert chart_xml.count('<a:noFill') >= 2
