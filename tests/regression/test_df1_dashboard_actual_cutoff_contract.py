from __future__ import annotations

from datetime import datetime

import pytest
from openpyxl import Workbook

from progress_studio.domain.main_dataset import MainDataset, MainPeriod
from progress_studio.infrastructure.excel.live_dashboard_workbook import _build_live_data_sheet


def _dataset() -> MainDataset:
    dates = (
        datetime(2026, 8, 7),
        datetime(2026, 8, 14),
        datetime(2026, 8, 21),
        datetime(2026, 8, 28),
        datetime(2026, 9, 4),
        datetime(2026, 9, 18),
    )
    return MainDataset(
        workbook_name="df1.xlsx",
        header_row=1,
        headers=(),
        periods=tuple(
            MainPeriod(column=20 + index, key=f"W{index + 1}", reporting_date=value)
            for index, value in enumerate(dates)
        ),
        rows=(),
    )


def _workbook() -> Workbook:
    workbook = Workbook()
    workbook.active.title = "Dashboard"
    workbook["Dashboard"]["G5"] = "Weekly"
    workbook["Dashboard"]["K5"] = datetime(2026, 9, 18)

    progress = workbook.create_sheet("progress")
    progress.append(["Date", "Plan", "Actual"])
    for values in (
        (datetime(2026, 8, 7), 0.0031, 0.0024),
        (datetime(2026, 8, 14), 0.0050, 0.0038),
        (datetime(2026, 8, 21), 0.0069, ""),
        (datetime(2026, 8, 28), 0.0098, ""),
        (datetime(2026, 9, 4), 0.0117, ""),
        (datetime(2026, 9, 18), 0.0202, ""),
    ):
        progress.append(values)
    return workbook


@pytest.mark.smoke
@pytest.mark.regression
def test_df1_weekly_raw_actual_carries_latest_value_to_reporting_cutoff():
    workbook = _workbook()
    _build_live_data_sheet(workbook, _dataset(), object())

    data = workbook["Dashboard_Data"]
    formula = data["L7"].value

    assert "LOOKUP(2" in formula
    assert "$J$2:$J$7<=G7" in formula
    assert '$C$2:$C$7<>""' in formula
    assert "$C$2:$C$7" in formula
    assert "Dashboard!$K$5" not in formula
    assert "G7>Dashboard!$K$5" in data["I7"].value
    assert "L7" in data["I7"].value


@pytest.mark.smoke
@pytest.mark.regression
def test_df1_monthly_raw_actual_uses_weekly_history_not_exact_month_end():
    workbook = _workbook()
    workbook["Dashboard"]["G5"] = "Monthly"
    _build_live_data_sheet(workbook, _dataset(), object())

    data = workbook["Dashboard_Data"]
    formula = data["L3"].value

    assert "LOOKUP(2" in formula
    assert "$J$2:$J$7<=G3" in formula
    assert "$C$2:$C$7" in formula
    assert "F3" not in formula


@pytest.mark.smoke
@pytest.mark.regression
def test_df1_marker_ownership_remains_on_raw_actual_helper():
    workbook = _workbook()
    _build_live_data_sheet(workbook, _dataset(), object())

    data = workbook["Dashboard_Data"]
    assert "SUMIFS($L$2:$L$7" in data["O2"].value
    assert "SUMIFS($H$2:$H$7" in data["N2"].value
