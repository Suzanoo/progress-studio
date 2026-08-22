
from __future__ import annotations

from pathlib import Path


def test_dashboard_monthly_curve_prefers_progress_contract() -> None:
    source = Path(
        "progress_studio/infrastructure/excel/live_dashboard_workbook.py"
    ).read_text(encoding="utf-8")
    block = source[source.index("def _build_live_data_sheet"):source.index("def _kpi_box")]
    assert 'workbook["progress"]' in block
    assert "month_last_rows" in block
    assert "main_monthly" not in block
    assert 'Dashboard!$K$5' in block
    assert 'NA()' in block


def test_activity_pair_filter_keeps_plan_status_value_but_hides_it() -> None:
    source = Path(
        "progress_studio/infrastructure/excel/live_dashboard_workbook.py"
    ).read_text(encoding="utf-8")
    block = source[source.index("def _write_activity_section"):source.index("def build_live_dashboard")]
    assert '"Not Due"' in block
    assert '"No Progress"' in block
    assert 'ws.auto_filter.ref = f"P38:P{last_activity_row}"' in block
    assert 'ws[f"P{row}"].font = Font(name=_FONT, color=base_fill, size=9)' in block


def test_monthly_scurve_four_rows_receive_main_palette() -> None:
    source = Path(
        "progress_studio/infrastructure/excel/live_monthly_workbook.py"
    ).read_text(encoding="utf-8")
    assert '"P": SCURVE_PLAN_FILL' in source
    assert '"AP": WBS_PLAN_FILL' in source
    assert '"A": SCURVE_ACTUAL_FILL' in source
    assert '"AA": WBS_ACTUAL_FILL' in source
    assert 'row_type != "s-curve"' in source


def test_plan_curve_full_actual_curve_cutoff_contract_is_explicit() -> None:
    source = Path(
        "progress_studio/infrastructure/excel/live_dashboard_workbook.py"
    ).read_text(encoding="utf-8")
    assert "progress is the sole curve calculation contract" in source
    assert "renderer-only cutoff mask" in source
    assert "Actual history" in source
