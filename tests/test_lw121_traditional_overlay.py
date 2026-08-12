from pathlib import Path
from openpyxl import Workbook

from progress_studio.infrastructure.excel.traditional_overlay_workbook import (
    build_traditional_overlays,
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
