from datetime import datetime

import pytest

from progress_studio.domain.main_dataset import MainDataset, MainPeriod, MainRow
from progress_studio.domain.mapping_models import AllocationRecord, BOQRow
from progress_studio.services.earned_value_deriver import (
    EarnedValueDeriver,
    EarnedValueInputError,
)


def _row(*, pa: str, activity_id: str, amount: float | None, values: tuple[float | None, ...], wbs: str = "1") -> MainRow:
    return MainRow(
        row_number=1,
        row_type="Activity" if pa == "P" else "",
        pa=pa,
        wbs=wbs,
        description=f"Activity {activity_id}" if pa == "P" else "",
        activity_id=activity_id,
        outline_level=2,
        plan_start=None,
        plan_finish=None,
        amount=amount,
        percent_complete=None,
        period_values=tuple((10 + i, value) for i, value in enumerate(values)),
    )


def _dataset(rows: tuple[MainRow, ...]) -> MainDataset:
    periods = (
        MainPeriod(10, "W1", datetime(2026, 8, 7)),
        MainPeriod(11, "W2", datetime(2026, 8, 14)),
        MainPeriod(12, "W3", datetime(2026, 8, 21)),
        MainPeriod(13, "W4", datetime(2026, 8, 28)),
    )
    return MainDataset("test.xlsx", 1, (), periods, rows)


@pytest.mark.unit
def test_project_ev_uses_weekly_increment_then_cumulative_progress() -> None:
    dataset = _dataset((
        _row(pa="P", activity_id="A1", amount=100.0, values=(0.10, 0.20, 0.30, 0.40)),
        _row(pa="A", activity_id="A1", amount=None, values=(0.05, 0.15, 0.20, 0.10)),
    ))
    boq = (BOQRow("B1", "BOQ", 1, "", "", "", "Item", 100.0, "BOQ-001"),)
    allocations = (AllocationRecord("B1", "A1", 100.0),)

    result = EarnedValueDeriver().derive(
        dataset, boq, allocations, cutoff_date=datetime(2026, 8, 21)
    )

    assert result.project_bac == pytest.approx(100.0)
    assert [p.planned_value for p in result.project_points] == pytest.approx([10, 30, 60, 100])
    assert [p.earned_value for p in result.project_points[:3]] == pytest.approx([5, 20, 40])
    assert result.project_points[3].earned_value is None
    assert result.project_points[2].schedule_variance == pytest.approx(-20.0)
    assert result.project_points[2].schedule_performance_index == pytest.approx(40 / 60)


@pytest.mark.unit
def test_boq_ev_reverse_aggregates_using_each_allocation_amount() -> None:
    dataset = _dataset((
        _row(pa="P", activity_id="A100", amount=20.0, values=(0.25, 0.25, 0.25, 0.25)),
        _row(pa="A", activity_id="A100", amount=None, values=(0.25, 0.25, 0.25, 0.25)),
        _row(pa="P", activity_id="A200", amount=30.0, values=(0.20, 0.20, 0.20, 0.20)),
        _row(pa="A", activity_id="A200", amount=None, values=(0.10, 0.10, 0.20, 0.20)),
        _row(pa="P", activity_id="A300", amount=50.0, values=(0.10, 0.10, 0.10, 0.10)),
        _row(pa="A", activity_id="A300", amount=None, values=(0.05, 0.05, 0.05, 0.05)),
    ))
    boq = (BOQRow("STEEL", "BOQ", 8, "", "", "", "Reinforcement", 100.0, "BOQ-STEEL"),)
    allocations = (
        AllocationRecord("STEEL", "A100", 20.0),
        AllocationRecord("STEEL", "A200", 30.0),
        AllocationRecord("STEEL", "A300", 50.0),
    )

    result = EarnedValueDeriver().derive(
        dataset, boq, allocations, cutoff_date=datetime(2026, 8, 28)
    )
    point = result.boq_items[0].points[-1]

    # At W4: plan = 20*1.0 + 30*0.8 + 50*0.4 = 64
    #         EV   = 20*1.0 + 30*0.6 + 50*0.2 = 48
    assert point.planned_value == pytest.approx(64.0)
    assert point.earned_value == pytest.approx(48.0)
    assert point.schedule_variance == pytest.approx(-16.0)
    assert point.schedule_performance_index == pytest.approx(0.75)


@pytest.mark.unit
def test_boq_allocation_not_100_percent_is_hard_stop() -> None:
    dataset = _dataset((
        _row(pa="P", activity_id="A1", amount=80.0, values=(0.1, 0.1, 0.1, 0.1)),
        _row(pa="A", activity_id="A1", amount=None, values=(0.1, 0.1, 0.1, 0.1)),
    ))
    boq = (BOQRow("B1", "BOQ", 1, "", "", "", "Item", 100.0, "BOQ-001"),)

    with pytest.raises(EarnedValueInputError, match=r"BOQ-001: 80\.00%"):
        EarnedValueDeriver().derive(
            dataset,
            boq,
            (AllocationRecord("B1", "A1", 80.0),),
            cutoff_date=datetime(2026, 8, 28),
        )


@pytest.mark.unit
def test_allocation_tolerance_accepts_rounding_noise() -> None:
    dataset = _dataset((
        _row(pa="P", activity_id="A1", amount=100.0, values=(0.1, 0.1, 0.1, 0.1)),
        _row(pa="A", activity_id="A1", amount=None, values=(0.1, 0.1, 0.1, 0.1)),
    ))
    boq = (BOQRow("B1", "BOQ", 1, "", "", "", "Item", 100.0),)

    result = EarnedValueDeriver().derive(
        dataset,
        boq,
        (AllocationRecord("B1", "A1", 99.995),),
        cutoff_date=datetime(2026, 8, 28),
    )
    assert result.boq_items[0].bac == pytest.approx(100.0)


@pytest.mark.unit
def test_unknown_activity_allocation_is_rejected() -> None:
    dataset = _dataset((
        _row(pa="P", activity_id="A1", amount=100.0, values=(0.1, 0.1, 0.1, 0.1)),
        _row(pa="A", activity_id="A1", amount=None, values=(0.1, 0.1, 0.1, 0.1)),
    ))
    boq = (BOQRow("B1", "BOQ", 1, "", "", "", "Item", 100.0),)

    with pytest.raises(EarnedValueInputError, match="unknown Activity ID"):
        EarnedValueDeriver().derive(
            dataset,
            boq,
            (AllocationRecord("B1", "MISSING", 100.0),),
            cutoff_date=datetime(2026, 8, 28),
        )

@pytest.mark.unit
def test_margin_periods_outside_plan_reporting_range_are_excluded() -> None:
    periods = (
        MainPeriod(9, "MARGIN-L", datetime(2026, 7, 31)),
        MainPeriod(10, "W1", datetime(2026, 8, 7)),
        MainPeriod(11, "W2", datetime(2026, 8, 14)),
        MainPeriod(12, "MARGIN-R", datetime(2026, 8, 21)),
    )
    plan = MainRow(
        row_number=1, row_type="Activity", pa="P", wbs="1", description="A1",
        activity_id="A1", outline_level=2, plan_start=None, plan_finish=None,
        amount=100.0, percent_complete=None,
        period_values=((9, None), (10, 0.5), (11, 0.5), (12, None)),
    )
    actual = MainRow(
        row_number=2, row_type="", pa="A", wbs="", description="",
        activity_id="A1", outline_level=None, plan_start=None, plan_finish=None,
        amount=None, percent_complete=None,
        period_values=((9, None), (10, 0.4), (11, 0.4), (12, None)),
    )
    dataset = MainDataset("test.xlsx", 1, (), periods, (plan, actual))
    boq = (BOQRow("B1", "BOQ", 1, "", "", "", "Item", 100.0),)

    result = EarnedValueDeriver().derive(
        dataset, boq, (AllocationRecord("B1", "A1", 100.0),),
        cutoff_date=datetime(2026, 8, 14),
    )

    assert [point.period_key for point in result.project_points] == ["W1", "W2"]
    assert [point.planned_value for point in result.project_points] == pytest.approx([50, 100])
