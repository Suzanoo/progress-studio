from __future__ import annotations

from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook, load_workbook

from progress_studio.domain.earned_value import EarnedValuePoint, EarnedValueResult
from progress_studio.services.earned_value_rebuild_service import (
    EarnedValueRebuildError,
    EarnedValueRebuildService,
)


CUTOFF = datetime(2026, 8, 28)


class StubInputReader:
    def __init__(self) -> None:
        self.boq_rows = (SimpleNamespace(key="B1"), SimpleNamespace(key="B2"))
        self.allocations = (
            SimpleNamespace(boq_key="B1", activity_id="A1", share_percent=100.0),
            SimpleNamespace(boq_key="B2", activity_id="A2", share_percent=100.0),
        )

    def read(self, path: Path):
        return SimpleNamespace(
            workbook=path,
            boq_rows=self.boq_rows,
            allocations=self.allocations,
        )


class StubMainReader:
    def read_main_dataset(self, path: Path):
        return SimpleNamespace(
            activities=(SimpleNamespace(activity_id="A1"), SimpleNamespace(activity_id="A2"))
        )


class StubDeriver:
    def derive(self, dataset, boq_rows, allocations, *, cutoff_date):
        assert cutoff_date == CUTOFF
        return EarnedValueResult(
            cutoff_date=cutoff_date,
            project_bac=1000.0,
            project_points=(
                EarnedValuePoint(
                    period_key="W1",
                    reporting_date=CUTOFF,
                    planned_value=600.0,
                    earned_value=500.0,
                    schedule_variance=-100.0,
                    schedule_performance_index=5.0 / 6.0,
                ),
            ),
            activities=(),
            boq_items=(),
        )


def _source_workbook(tmp_path: Path, *, cutoff=CUTOFF, earned_value=False) -> Path:
    workbook = Workbook()
    main = workbook.active
    main.title = "main"
    main["A1"] = "user main data"
    dashboard = workbook.create_sheet("Dashboard")
    dashboard["K5"] = cutoff
    if earned_value:
        ws = workbook.create_sheet("Earned Value")
        ws["A1"] = "old"
    path = tmp_path / "source.xlsx"
    workbook.save(path)
    workbook.close()
    return path


def _service(renderer=None) -> EarnedValueRebuildService:
    kwargs = dict(
        input_reader=StubInputReader(),
        main_reader=StubMainReader(),
        deriver=StubDeriver(),
    )
    if renderer is not None:
        kwargs["renderer"] = renderer
    return EarnedValueRebuildService(**kwargs)


@pytest.mark.unit
def test_ev_rebuild_analysis_reports_readiness_and_existing_sheet(tmp_path: Path) -> None:
    source = _source_workbook(tmp_path, earned_value=True)

    analysis = _service().analyze(source)

    assert analysis.cutoff_date == CUTOFF
    assert analysis.activity_count == 2
    assert analysis.boq_count == 2
    assert analysis.allocation_count == 2
    assert analysis.project_bac == pytest.approx(1000.0)
    assert analysis.existing_earned_value_sheet is True


@pytest.mark.unit
def test_ev_rebuild_generate_preserves_existing_sheets_and_refreshes_only_ev(
    tmp_path: Path,
) -> None:
    source = _source_workbook(tmp_path, earned_value=True)
    output = tmp_path / "output.xlsx"

    def renderer(workbook, result):
        if "Earned Value" in workbook.sheetnames:
            del workbook["Earned Value"]
        ws = workbook.create_sheet("Earned Value")
        ws["A1"] = "EV-4"
        ws["B1"] = result.project_bac

    result = _service(renderer=renderer).generate(source, output)

    assert result.refreshed_sheet == "Earned Value"
    assert source.exists()
    workbook = load_workbook(output, data_only=False)
    try:
        assert workbook["main"]["A1"].value == "user main data"
        assert workbook["Dashboard"]["K5"].value == CUTOFF
        assert workbook["Earned Value"]["A1"].value == "EV-4"
        assert workbook["Earned Value"]["B1"].value == pytest.approx(1000.0)
    finally:
        workbook.close()


@pytest.mark.unit
def test_ev_rebuild_uses_main_cutoff_as_fallback(tmp_path: Path) -> None:
    workbook = Workbook()
    main = workbook.active
    main.title = "main"
    main["L5"] = "Cutoff Date"
    main["M5"] = CUTOFF
    path = tmp_path / "fallback.xlsx"
    workbook.save(path)
    workbook.close()

    analysis = _service().analyze(path)

    assert analysis.cutoff_date == CUTOFF


@pytest.mark.unit
def test_ev_rebuild_hard_stops_without_cutoff(tmp_path: Path) -> None:
    source = _source_workbook(tmp_path, cutoff=None)

    with pytest.raises(EarnedValueRebuildError, match="requires a reporting cutoff date"):
        _service().analyze(source)


@pytest.mark.unit
def test_ev_rebuild_refuses_in_place_output(tmp_path: Path) -> None:
    source = _source_workbook(tmp_path)

    with pytest.raises(EarnedValueRebuildError, match="must be a new workbook path"):
        _service().generate(source, source)
