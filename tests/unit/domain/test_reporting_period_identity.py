from __future__ import annotations

from datetime import date, datetime, timedelta

import pytest

from progress_studio.domain.reporting_period_identity import (
    ReportingPeriodWindow,
    number_reporting_periods,
)


def _weekly_window(week_ending: date) -> ReportingPeriodWindow:
    return ReportingPeriodWindow(week_ending - timedelta(days=6), week_ending)


def _labels(result):
    return [item.label for item in result]


def test_rn2_weekly_numbering_without_margin_is_contiguous() -> None:
    periods = tuple(
        _weekly_window(date(2026, 2, 27) + timedelta(days=7 * offset))
        for offset in range(3)
    )

    result = number_reporting_periods(
        periods,
        date(2026, 2, 23),
        date(2026, 3, 13),
        prefix="W",
    )

    assert _labels(result) == ["W1", "W2", "W3"]
    assert [item.sequence for item in result] == [1, 2, 3]


def test_rn2_pre_margin_has_no_reporting_identity() -> None:
    endings = [
        date(2026, 1, 30),
        date(2026, 2, 6),
        date(2026, 2, 13),
        date(2026, 2, 20),
        date(2026, 2, 27),
        date(2026, 3, 6),
    ]

    result = number_reporting_periods(
        tuple(_weekly_window(value) for value in endings),
        date(2026, 2, 23),
        date(2026, 3, 6),
        prefix="W",
    )

    assert _labels(result) == [None, None, None, None, "W1", "W2"]


def test_rn2_post_margin_has_no_reporting_identity() -> None:
    endings = [
        date(2026, 2, 27),
        date(2026, 3, 6),
        date(2026, 3, 13),
        date(2026, 3, 20),
    ]

    result = number_reporting_periods(
        tuple(_weekly_window(value) for value in endings),
        date(2026, 2, 23),
        date(2026, 3, 6),
        prefix="W",
    )

    assert _labels(result) == ["W1", "W2", None, None]


def test_rn2_both_margins_do_not_create_numbering_gaps() -> None:
    endings = [
        date(2026, 2, 13),
        date(2026, 2, 20),
        date(2026, 2, 27),
        date(2026, 3, 6),
        date(2026, 3, 13),
        date(2026, 3, 20),
    ]

    result = number_reporting_periods(
        tuple(_weekly_window(value) for value in endings),
        date(2026, 2, 23),
        date(2026, 3, 10),
        prefix="W",
    )

    assert _labels(result) == [None, None, "W1", "W2", "W3", None]
    assert [item.sequence for item in result if item.is_reporting] == [1, 2, 3]


def test_rn2_project_start_inside_period_makes_that_period_w1() -> None:
    periods = (
        ReportingPeriodWindow(date(2026, 2, 21), date(2026, 2, 27)),
        ReportingPeriodWindow(date(2026, 2, 28), date(2026, 3, 6)),
    )

    result = number_reporting_periods(
        periods,
        date(2026, 2, 23),
        date(2026, 3, 1),
        prefix="W",
    )

    assert _labels(result) == ["W1", "W2"]


def test_rn2_project_finish_inside_period_keeps_final_period() -> None:
    periods = (
        ReportingPeriodWindow(date(2027, 5, 22), date(2027, 5, 28)),
        ReportingPeriodWindow(date(2027, 5, 29), date(2027, 6, 4)),
        ReportingPeriodWindow(date(2027, 6, 5), date(2027, 6, 11)),
    )

    result = number_reporting_periods(
        periods,
        date(2027, 5, 22),
        date(2027, 6, 1),
        prefix="W",
    )

    assert _labels(result) == ["W1", "W2", None]


def test_rn2_one_period_project_gets_single_identity() -> None:
    periods = (
        ReportingPeriodWindow(date(2026, 8, 1), date(2026, 8, 31)),
    )

    result = number_reporting_periods(
        periods,
        date(2026, 8, 15),
        date(2026, 8, 15),
        prefix="M",
    )

    assert _labels(result) == ["M1"]


def test_rn2_monthly_periods_use_same_contract_as_weekly() -> None:
    periods = (
        ReportingPeriodWindow(date(2026, 1, 1), date(2026, 1, 31)),
        ReportingPeriodWindow(date(2026, 2, 1), date(2026, 2, 28)),
        ReportingPeriodWindow(date(2026, 3, 1), date(2026, 3, 31)),
        ReportingPeriodWindow(date(2026, 4, 1), date(2026, 4, 30)),
        ReportingPeriodWindow(date(2026, 5, 1), date(2026, 5, 31)),
    )

    result = number_reporting_periods(
        periods,
        date(2026, 2, 23),
        date(2026, 4, 10),
        prefix="M",
    )

    assert _labels(result) == [None, "M1", "M2", "M3", None]


def test_rn2_datetime_boundaries_are_normalized_to_dates() -> None:
    periods = (
        ReportingPeriodWindow(date(2026, 2, 21), date(2026, 2, 27)),
    )

    result = number_reporting_periods(
        periods,
        datetime(2026, 2, 23, 8, 0),
        datetime(2026, 2, 23, 17, 0),
        prefix="W",
    )

    assert result[0].label == "W1"


def test_rn2_invalid_project_boundary_is_rejected() -> None:
    with pytest.raises(ValueError, match="project_finish"):
        number_reporting_periods(
            (),
            date(2026, 3, 2),
            date(2026, 3, 1),
            prefix="W",
        )


def test_rn2_invalid_period_window_is_rejected() -> None:
    with pytest.raises(ValueError, match="period end"):
        ReportingPeriodWindow(date(2026, 3, 2), date(2026, 3, 1))


def test_rn2_blank_prefix_is_rejected() -> None:
    with pytest.raises(ValueError, match="prefix"):
        number_reporting_periods(
            (),
            date(2026, 3, 1),
            date(2026, 3, 2),
            prefix="   ",
        )


def test_rn2_is_pure_and_preserves_input_order() -> None:
    periods = (
        ReportingPeriodWindow(date(2026, 2, 1), date(2026, 2, 7)),
        ReportingPeriodWindow(date(2026, 2, 8), date(2026, 2, 14)),
        ReportingPeriodWindow(date(2026, 2, 15), date(2026, 2, 21)),
    )
    original = periods

    result = number_reporting_periods(
        periods,
        date(2026, 2, 8),
        date(2026, 2, 14),
        prefix="W",
    )

    assert periods == original
    assert [item.source_index for item in result] == [0, 1, 2]
    assert [item.window for item in result] == list(periods)
    assert _labels(result) == [None, "W1", None]
