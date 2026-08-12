"""LW-13 Dashboard freeze regression gate.

These tests intentionally duplicate a small set of high-value invariants from
older milestone tests.  The older tests explain each feature; this file is the
single release gate for the accepted Dashboard V1 behavior.
"""

from pathlib import Path

from progress_studio.domain.dashboard_freeze import DASHBOARD_V1
from progress_studio.infrastructure.excel import traditional_overlay_workbook as overlay


ROOT = Path("progress_studio")


def test_dashboard_v1_contract_is_versioned_and_immutable():
    assert DASHBOARD_V1.version == "dashboard-v1-stable"
    assert DASHBOARD_V1.dashboard_view_cell == "G5"
    assert DASHBOARD_V1.dashboard_cutoff_cell == "K5"
    assert DASHBOARD_V1.traditional_cutoff_column == "M"
    assert DASHBOARD_V1.weekly_cutoff_format == "dd/mm/yyyy"
    assert DASHBOARD_V1.monthly_cutoff_format == "mmm yyyy"
    assert DASHBOARD_V1.plan_uses_full_timeline is True
    assert DASHBOARD_V1.actual_is_cutoff_masked is True
    assert DASHBOARD_V1.traditional_cutoffs_are_independent is True
    assert DASHBOARD_V1.traditional_overlay_has_legend is False
    assert DASHBOARD_V1.cutoff_line_has_label is True
    assert DASHBOARD_V1.cutoff_label_font_points == 10


def test_dashboard_v1_live_dashboard_control_and_curve_contract():
    text = (ROOT / "infrastructure/excel/live_dashboard_workbook.py").read_text()

    # Stable controls.
    assert 'ws["G5"] = "Weekly"' in text
    assert 'ws["K5"] = cutoff_date' in text
    assert 'view_validation.add(ws["G5"])' in text
    assert 'cutoff_validation.add(ws["K5"])' in text

    # Plan remains the full selected series; cutoff masks Actual only.
    assert "min_col=8, max_col=9" in text
    assert 'IF(G{row}>Dashboard!$K$5,NA()' in text


def test_dashboard_v1_traditional_overlay_contract():
    text = (ROOT / "infrastructure/excel/traditional_overlay_workbook.py").read_text()

    # Weekly and Monthly own independent cutoff cells in column M.
    assert "label_col = 12  # L" in text
    assert "value_col = 13  # M" in text
    assert 'display_format="dd/mm/yyyy"' in text
    assert 'display_format="mmm yyyy"' in text
    assert "weekly_cutoff_ref=weekly_cutoff_ref" in text
    assert "monthly_cutoff_ref=monthly_cutoff_ref" in text

    # Actual is cutoff-masked; Plan is not.  Red cutoff line is labelled,
    # legible, and the overlay intentionally has no legend.
    assert 'IF(A{row}>{weekly_cutoff_ref},NA()' in text
    assert 'IF(D{row}>{monthly_cutoff_ref},NA()' in text
    assert 'SeriesLabel(v="Cutoff")' in text
    assert "chart.legend = None" in text
    assert overlay.CUTOFF_LABEL_FONT_SIZE == 1000


def test_dashboard_v1_rebuild_pipeline_still_builds_dashboard_and_overlays():
    text = (ROOT / "services/rebuild_service.py").read_text()
    assert "build_live_dashboard(" in text
    assert "build_traditional_overlays(wb, dataset)" in text
