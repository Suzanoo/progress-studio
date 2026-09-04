from datetime import datetime

import pytest

from progress_studio.domain.main_dataset import MainDataset, MainPeriod, MainRow
from progress_studio.services.payment_breakdown_adapter import (
    MainDatasetPaymentBreakdownAdapter,
)


def _row(
    row_number,
    activity_id,
    name,
    amount,
    profile,
    *,
    wbs="3.2.5.1",
):
    return MainRow(
        row_number=row_number,
        row_type="Activity",
        pa="P",
        wbs=wbs,
        description=name,
        activity_id=activity_id,
        outline_level=4,
        plan_start=datetime(2026, 1, 1),
        plan_finish=datetime(2026, 1, 31),
        amount=amount,
        percent_complete=0.0,
        period_values=tuple((6 + i, value) for i, value in enumerate(profile)),
    )


def _dataset(rows):
    periods = tuple(
        MainPeriod(
            column=6 + i,
            key=f"W{i + 1}",
            reporting_date=datetime(2026, 1, 2 + i * 7),
        )
        for i in range(3)
    )
    return MainDataset(
        workbook_name="test.xlsx",
        header_row=4,
        headers=(),
        periods=periods,
        rows=tuple(rows),
    )


def test_pb3_adapter_derives_first_and_second_fixed_as_exact_name_groups():
    rows = [
        _row(10, "A1", "First Fixed", 100.0, (0.5, 0.5, 0.0), wbs="3.2"),
        _row(11, "A2", "First Fixed", 300.0, (0.0, 0.5, 0.5), wbs="3.3"),
        _row(12, "B1", "Second Fixed", 200.0, (0.25, 0.5, 0.25), wbs="3.2"),
        _row(13, "B2", "Second Fixed", 200.0, (0.0, 0.25, 0.75), wbs="3.3"),
        _row(14, "C1", "Singleton", 50.0, (0.0, 0.0, 1.0)),
    ]

    result = MainDatasetPaymentBreakdownAdapter().derive(_dataset(rows))

    assert [item.activity_name for item in result.activities] == [
        "First Fixed",
        "Second Fixed",
    ]
    assert result.eligible_source_count == 5
    assert result.skipped_activity_ids == ()

    first = result.activities[0]
    assert first.activity_ids == ("A1", "A2")
    assert first.total_amount == pytest.approx(400.0)
    assert first.period_progress == pytest.approx((0.125, 0.5, 0.375))
    assert first.cumulative_progress == pytest.approx((0.125, 0.625, 1.0))


def test_pb3_adapter_skips_nonpositive_or_incomplete_source_rows_before_grouping():
    rows = [
        _row(10, "A1", "First Fixed", 100.0, (0.5, 0.5, 0.0)),
        _row(11, "A2", "First Fixed", 0.0, (0.0, 0.0, 1.0)),
        _row(12, "A3", "First Fixed", 100.0, (0.2, 0.2, 0.2)),
        _row(13, "A4", "First Fixed", 100.0, (0.0, 0.5, 0.5)),
    ]

    result = MainDatasetPaymentBreakdownAdapter().derive(_dataset(rows))

    assert len(result.activities) == 1
    assert result.activities[0].activity_ids == ("A1", "A4")
    assert result.eligible_source_count == 2
    assert result.skipped_activity_ids == ("A2", "A3")


def test_pb3_adapter_never_contains_matches_distinct_activity_names():
    rows = [
        _row(10, "U1", "P3.1-P3.3 | U-Glass", 100.0, (1.0, 0.0, 0.0)),
        _row(11, "U2", "P3.2 U-Glass Wall", 100.0, (0.0, 1.0, 0.0)),
        _row(12, "U3", "P3.2, 3.2 U-Glass Wall", 100.0, (0.0, 0.0, 1.0)),
    ]

    result = MainDatasetPaymentBreakdownAdapter().derive(_dataset(rows))

    assert result.activities == ()
