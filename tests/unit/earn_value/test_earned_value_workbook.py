from __future__ import annotations

from datetime import datetime

from openpyxl import Workbook

from progress_studio.domain.earned_value import EarnedValuePoint, EarnedValueResult
from progress_studio.infrastructure.excel.earned_value_workbook import (
    EARNED_VALUE_SHEET,
    render_earned_value_sheet,
)


def _point(
    key: str,
    reporting_date: datetime,
    pv: float | None,
    ev: float | None,
) -> EarnedValuePoint:
    return EarnedValuePoint(
        period_key=key,
        reporting_date=reporting_date,
        planned_value=pv,
        earned_value=ev,
        schedule_variance=None if pv is None or ev is None else ev - pv,
        schedule_performance_index=(
            None if pv in (None, 0.0) or ev is None else ev / pv
        ),
    )


def _result() -> EarnedValueResult:
    return EarnedValueResult(
        cutoff_date=datetime(2026, 2, 13),
        project_bac=10_000_000.0,
        project_points=(
            _point("W1", datetime(2026, 1, 30), 2_000_000.0, 1_500_000.0),
            _point("W2", datetime(2026, 2, 6), 4_000_000.0, 3_000_000.0),
            _point("W3", datetime(2026, 2, 13), 6_000_000.0, 4_500_000.0),
            _point("W4", datetime(2026, 2, 20), 8_000_000.0, None),
        ),
        activities=(),
        boq_items=(),
    )


def test_render_creates_project_ev_sheet_with_cutoff_kpis() -> None:
    workbook = Workbook()

    render_earned_value_sheet(workbook, _result())

    assert EARNED_VALUE_SHEET in workbook.sheetnames
    ws = workbook[EARNED_VALUE_SHEET]
    assert ws["A1"].value == "EARNED VALUE"
    assert ws["B4"].value == datetime(2026, 2, 13)

    assert ws["A6"].value == "BAC"
    assert ws["A7"].value == 10_000_000.0
    assert ws["C6"].value == "PV"
    assert ws["C7"].value == 6_000_000.0
    assert ws["E6"].value == "EV"
    assert ws["E7"].value == 4_500_000.0
    assert ws["G6"].value == "SV"
    assert ws["G7"].value == -1_500_000.0
    assert ws["I6"].value == "SPI"
    assert ws["I7"].value == 0.75


def test_render_writes_full_pv_horizon_and_stops_ev_after_cutoff() -> None:
    workbook = Workbook()

    render_earned_value_sheet(workbook, _result())

    ws = workbook[EARNED_VALUE_SHEET]
    assert ws["A12"].value == "W1"
    assert ws["C15"].value == 8_000_000.0
    assert ws["D15"].value is None
    assert len(ws._charts) == 1
    assert ws._charts[0].title is not None


def test_render_replaces_only_earned_value_sheet() -> None:
    workbook = Workbook()
    keep = workbook.active
    keep.title = "main"
    old = workbook.create_sheet(EARNED_VALUE_SHEET)
    old["A1"] = "old"

    render_earned_value_sheet(workbook, _result())

    assert "main" in workbook.sheetnames
    assert workbook["main"]["A1"].value is None
    assert workbook[EARNED_VALUE_SHEET]["A1"].value == "EARNED VALUE"
    assert workbook.sheetnames.count(EARNED_VALUE_SHEET) == 1
