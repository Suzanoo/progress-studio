from datetime import datetime

import pytest

from progress_studio.domain.main_dataset import MainDataset, MainPeriod, MainRow
from progress_studio.domain.mapping_models import AllocationRecord, BOQRow
from progress_studio.services.earned_value_deriver import EarnedValueDeriver


def _row(*, pa: str, activity_id: str, amount: float | None, values: tuple[float | None, ...]) -> MainRow:
    return MainRow(
        row_number=1,
        row_type="Activity" if pa == "P" else "",
        pa=pa,
        wbs="1",
        description=f"Activity {activity_id}" if pa == "P" else "",
        activity_id=activity_id,
        outline_level=2,
        plan_start=None,
        plan_finish=None,
        amount=amount,
        percent_complete=None,
        period_values=tuple((10 + i, value) for i, value in enumerate(values)),
    )


@pytest.mark.unit
def test_project_ev_uses_zero_baseline_before_first_actual() -> None:
    periods = (
        MainPeriod(10, "W1", datetime(2026, 8, 7)),
        MainPeriod(11, "W2", datetime(2026, 8, 14)),
        MainPeriod(12, "W3", datetime(2026, 8, 21)),
        MainPeriod(13, "W4", datetime(2026, 8, 28)),
    )
    dataset = MainDataset(
        "test.xlsx",
        1,
        (),
        periods,
        (
            _row(pa="P", activity_id="A1", amount=100.0, values=(0.10, 0.20, 0.30, 0.40)),
            _row(pa="A", activity_id="A1", amount=None, values=(None, None, None, None)),
        ),
    )
    result = EarnedValueDeriver().derive(
        dataset,
        (BOQRow("B1", "BOQ", 1, "", "", "", "Item", 100.0, "BOQ-001"),),
        (AllocationRecord("B1", "A1", 100.0),),
        cutoff_date=datetime(2026, 8, 21),
    )

    assert [p.earned_value for p in result.project_points[:3]] == pytest.approx([0.0, 0.0, 0.0])
    assert result.project_points[3].earned_value is None
    assert result.project_points[2].schedule_variance == pytest.approx(-60.0)
    assert result.project_points[2].schedule_performance_index == pytest.approx(0.0)
