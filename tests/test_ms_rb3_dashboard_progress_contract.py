from __future__ import annotations

from datetime import date

from openpyxl import Workbook

from progress_studio.infrastructure.excel.dashboard_workbook import (
    DATA_SHEET,
    DASHBOARD_SHEET,
    build_dashboard,
)


def _dashboard_source() -> Workbook:
    wb = Workbook()
    progress = wb.active
    progress.title = "progress"
    progress.append(
        ["project_start", "project_finish", "week_start", "plan", "actual"]
    )
    # progress contract = 0..100 percent-points, including values below 1%.
    rows = [
        (date(2026, 3, 20), 0.31, None),
        (date(2026, 3, 27), 0.48, 0.20),
        (date(2026, 4, 3), 0.75, 0.30),
        (date(2026, 4, 10), 1.20, 0.45),
        (date(2026, 4, 17), 2.50, None),
    ]
    for week, plan, actual in rows:
        progress.append(
            [date(2026, 3, 1), date(2027, 5, 31), week, plan, actual]
        )

    table = wb.create_sheet("progress_table")
    table.append(
        ["WBS", "Activities", "Amount", "P/A", "%Progress", date(2026, 3, 20)]
    )
    table.append(["1", "Activity", 100.0, "P", 100.0, 0.31])
    table.append(["1", "Activity", 100.0, "A", 45.0, 0.20])
    return wb


def test_dashboard_adapter_treats_progress_as_percent_points() -> None:
    wb = _dashboard_source()
    build_dashboard(wb)

    data = wb[DATA_SHEET]
    assert data["B2"].value == 0.0031
    assert data["B3"].value == 0.0048
    assert data["B5"].value == 0.012
    assert data["C3"].value == 0.002


def test_monthly_curve_uses_last_cumulative_progress_value() -> None:
    wb = _dashboard_source()
    build_dashboard(wb)

    data = wb[DATA_SHEET]
    # March closes on 27-Mar.
    assert data["D2"].value == date(2026, 3, 27)
    assert data["E2"].value == 0.0048
    assert data["F2"].value == 0.002

    # April Plan closes on 17-Apr. Actual uses last populated cumulative
    # value (10-Apr), not SUM and not a LOOKUP formula.
    assert data["D3"].value == date(2026, 4, 17)
    assert data["E3"].value == 0.025
    assert abs(data["F3"].value - 0.0045) < 1e-12
    assert not isinstance(data["F3"].value, str)


def test_dashboard_still_uses_progress_table_for_activity_rows() -> None:
    wb = _dashboard_source()
    build_dashboard(wb)

    dashboard = wb[DASHBOARD_SHEET]
    assert dashboard["B39"].value == "='progress_table'!A2"
    assert dashboard["C39"].value == "='progress_table'!B2"
