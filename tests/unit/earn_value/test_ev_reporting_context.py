from __future__ import annotations

from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook, load_workbook

from progress_studio.domain.earned_value import EarnedValuePoint, EarnedValueResult
from progress_studio.infrastructure.excel.earned_value_workbook import (
    EARNED_VALUE_SHEET,
    EV_DATA_SHEET,
    render_earned_value_sheet,
)
from progress_studio.services.earned_value_rebuild_service import (
    EarnedValueRebuildError,
    EarnedValueRebuildService,
)


REPORTING_CUTOFF = datetime(2026, 8, 28)
STALE_EV_VIEW = datetime(2026, 4, 24)


class _InputReader:
    def read(self, path: Path):
        return SimpleNamespace(
            workbook=path,
            boq_rows=(SimpleNamespace(key="B1"),),
            allocations=(SimpleNamespace(
                boq_key="B1", activity_id="A1", share_percent=100.0
            ),),
        )


class _MainReader:
    def read_main_dataset(self, path: Path):
        return SimpleNamespace(
            activities=(SimpleNamespace(activity_id="A1"),),
            periods=(SimpleNamespace(column=18, reporting_date=REPORTING_CUTOFF),),
            rows=(),
        )


class _Deriver:
    def derive(self, dataset, boq_rows, allocations, *, cutoff_date):
        return EarnedValueResult(
            cutoff_date=cutoff_date,
            project_bac=1000.0,
            project_points=(
                EarnedValuePoint(
                    period_key="W1",
                    reporting_date=cutoff_date,
                    planned_value=600.0,
                    earned_value=500.0,
                    schedule_variance=-100.0,
                    schedule_performance_index=5.0 / 6.0,
                ),
            ),
            activities=(),
            boq_items=(),
        )


def _service() -> EarnedValueRebuildService:
    return EarnedValueRebuildService(
        input_reader=_InputReader(),
        main_reader=_MainReader(),
        deriver=_Deriver(),
    )


def _workbook_with_dashboard(tmp_path: Path, cutoff=REPORTING_CUTOFF) -> Path:
    workbook = Workbook()
    workbook.active.title = "main"
    dashboard = workbook.create_sheet("Dashboard")
    dashboard["K5"] = cutoff
    ev = workbook.create_sheet(EARNED_VALUE_SHEET)
    ev["M3"] = STALE_EV_VIEW
    path = tmp_path / "reporting-context.xlsx"
    workbook.save(path)
    workbook.close()
    return path


@pytest.mark.unit
def test_ev_refresh_preserves_existing_view_and_ignores_dashboard_cutoff(tmp_path: Path) -> None:
    path = _workbook_with_dashboard(tmp_path, cutoff=REPORTING_CUTOFF)
    assert EarnedValueRebuildService._read_existing_ev_view(path) == STALE_EV_VIEW
    assert _service().analyze(path).cutoff_date == STALE_EV_VIEW


@pytest.mark.unit
def test_ev_first_creation_uses_latest_canonical_monthly_view_not_dashboard_cutoff(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    workbook.active.title = "main"
    dashboard = workbook.create_sheet("Dashboard")
    dashboard["K5"] = STALE_EV_VIEW
    data = workbook.create_sheet("Dashboard_Data")
    data["K2"] = datetime(2026, 7, 31)
    data["K3"] = REPORTING_CUTOFF
    path = tmp_path / "canonical-view.xlsx"
    workbook.save(path)
    workbook.close()

    assert EarnedValueRebuildService._read_latest_canonical_view_date(path) == REPORTING_CUTOFF
    assert _service().analyze(path).cutoff_date == REPORTING_CUTOFF


@pytest.mark.unit
def test_ev_first_creation_falls_back_to_latest_main_reporting_date(tmp_path: Path) -> None:
    workbook = Workbook()
    workbook.active.title = "main"
    path = tmp_path / "main-period-fallback.xlsx"
    workbook.save(path)
    workbook.close()

    assert _service().analyze(path).cutoff_date == REPORTING_CUTOFF


def _point(key: str, when: datetime, pv: float, ev: float | None) -> EarnedValuePoint:
    return EarnedValuePoint(
        period_key=key,
        reporting_date=when,
        planned_value=pv,
        earned_value=ev,
        schedule_variance=None if ev is None else ev - pv,
        schedule_performance_index=None if ev is None or pv == 0 else ev / pv,
    )


def _ev_result() -> EarnedValueResult:
    return EarnedValueResult(
        cutoff_date=datetime(2026, 2, 13),
        project_bac=1000.0,
        project_points=(
            _point("W1", datetime(2026, 1, 30), 200.0, 150.0),
            _point("W2", datetime(2026, 2, 13), 400.0, 300.0),
            _point("W3", datetime(2026, 2, 27), 600.0, None),
            _point("W4", datetime(2026, 3, 27), 800.0, None),
        ),
        activities=(),
        boq_items=(),
    )


@pytest.mark.unit
def test_bf2_ev_dropdown_reuses_complete_dashboard_monthly_calendar() -> None:
    workbook = Workbook()
    dashboard_data = workbook.create_sheet("Dashboard_Data")
    dashboard_data["K1"] = "Monthly Cutoff"
    for row, value in enumerate((
        datetime(2026, 1, 30),
        datetime(2026, 2, 13),
        datetime(2026, 2, 27),
        datetime(2026, 3, 27),
    ), start=2):
        dashboard_data.cell(row, 11, value)

    render_earned_value_sheet(workbook, _ev_result(), include_chart=False)

    ev = workbook[EARNED_VALUE_SHEET]
    data = workbook[EV_DATA_SHEET]
    validation = ev.data_validations.dataValidation[0]
    assert validation.formula1 == '=INDIRECT("Dashboard_Data!$K$2:$K$5")'
    assert data["H1"].value is None


@pytest.mark.unit
def test_bf2_standalone_workbook_keeps_ev_data_fallback_calendar() -> None:
    workbook = Workbook()
    render_earned_value_sheet(workbook, _ev_result(), include_chart=False)

    ev = workbook[EARNED_VALUE_SHEET]
    data = workbook[EV_DATA_SHEET]
    validation = ev.data_validations.dataValidation[0]

    assert "EV_Data!$H$2:$H$" in validation.formula1
    assert data["H1"].value == "View Date Options"
    options = [
        data.cell(row, 8).value
        for row in range(2, data.max_row + 1)
        if isinstance(data.cell(row, 8).value, datetime)
    ]
    assert options == [datetime(2026, 1, 30), datetime(2026, 2, 27), datetime(2026, 3, 27)]

@pytest.mark.unit
def test_ev_view_seed_does_not_inspect_actual_progress_or_dashboard_cutoff(tmp_path: Path) -> None:
    path = _workbook_with_dashboard(tmp_path, cutoff=REPORTING_CUTOFF)
    # Existing EV view wins regardless of later Actual or Dashboard state.
    dataset = SimpleNamespace(
        periods=(SimpleNamespace(column=18, reporting_date=datetime(2026, 12, 25)),),
        rows=(SimpleNamespace(pa="A", activity_id="A1"),),
    )
    assert EarnedValueRebuildService._view_date_seed(path, dataset) == STALE_EV_VIEW
