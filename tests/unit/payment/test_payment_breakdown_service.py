import pytest

from progress_studio.services.payment_breakdown_service import (
    PaymentBreakdownError,
    PaymentBreakdownItem,
    PaymentBreakdownService,
    PaymentBreakdownSourceActivity,
)


UGLASS_AMOUNT = 3_778_840.27


def _seven_period_profile() -> tuple[float, ...]:
    return (1 / 14, 1 / 6, 1 / 6, 1 / 6, 1 / 6, 1 / 6, 2 / 21)


def _source(activity_id, name, amount, profile, wbs=None):
    return PaymentBreakdownSourceActivity(
        activity_id=activity_id,
        activity_name=name,
        amount=amount,
        period_progress=tuple(profile),
        wbs=wbs,
    )


def test_pb1_u_glass_exact_name_weighted_progress_contract():
    service = PaymentBreakdownService()
    profile = _seven_period_profile()
    items = (
        PaymentBreakdownItem("1", 1_446_031.38, (0.0,) * 11 + profile),
        PaymentBreakdownItem("M", 964_020.92, (0.0,) * 14 + profile),
        PaymentBreakdownItem("2", 394_550.05, profile),
        PaymentBreakdownItem("3", 974_237.92, (0.0,) * 4 + profile),
    )

    result = service.build_activity(
        main_activity_name="U-Glass",
        main_amount=UGLASS_AMOUNT,
        breakdown_activity_name="U-Glass",
        items=items,
    )

    assert result is not None
    assert result.activity_name == "U-Glass"
    assert result.breakdown_amount == pytest.approx(UGLASS_AMOUNT, abs=0.01)
    assert sum(result.period_progress) == pytest.approx(1.0)
    assert result.cumulative_progress[-1] == 1.0
    expected_first = (394_550.05 / UGLASS_AMOUNT) * (1 / 14)
    assert result.period_progress[0] == pytest.approx(expected_first)


def test_pb1_non_matching_activity_name_is_not_implemented():
    service = PaymentBreakdownService()
    result = service.build_activity(
        main_activity_name="U-Glass",
        main_amount=100.0,
        breakdown_activity_name="U Glass",
        items=(PaymentBreakdownItem("1", 100.0, (1.0,)),),
    )
    assert result is None


def test_pb1_leading_trailing_whitespace_is_harmless():
    service = PaymentBreakdownService()
    result = service.build_activity(
        main_activity_name=" U-Glass ",
        main_amount=100.0,
        breakdown_activity_name="U-Glass",
        items=(PaymentBreakdownItem("1", 100.0, (1.0,)),),
    )
    assert result is not None
    assert result.activity_name == "U-Glass"


def test_pb1_rejects_unbalanced_breakdown_amount():
    service = PaymentBreakdownService()
    with pytest.raises(PaymentBreakdownError, match="does not balance"):
        service.build_activity(
            main_activity_name="U-Glass",
            main_amount=100.0,
            breakdown_activity_name="U-Glass",
            items=(PaymentBreakdownItem("1", 99.0, (1.0,)),),
        )


def test_pb1_rejects_sub_activity_that_does_not_finish_at_100_percent():
    service = PaymentBreakdownService()
    with pytest.raises(PaymentBreakdownError, match="must total 100%"):
        service.build_activity(
            main_activity_name="U-Glass",
            main_amount=100.0,
            breakdown_activity_name="U-Glass",
            items=(PaymentBreakdownItem("1", 100.0, (0.25, 0.25)),),
        )


def test_pb2_first_fixed_keeps_five_source_progress_profiles_then_amount_weights_total():
    service = PaymentBreakdownService()
    rows = (
        _source("A1790", "First Fixed", 4_244_054.01, (0.25, 0.75), "3.2.5.1"),
        _source("A1960", "First Fixed", 1_522_296.53, (0.0, 1.0), "3.3.5.1"),
        _source("A2130", "First Fixed", 5_708_009.02, (0.50, 0.50), "3.5.5.1"),
        _source("A2280", "First Fixed", 5_202_860.60, (0.0, 1.0), "3.7.5.1"),
        _source("A2980", "First Fixed", 2_259_558.51, (1.0, 0.0), "3.9.5.1"),
    )

    result = service.derive_activity(rows)

    assert result.activity_name == "First Fixed"
    assert len(result.source_activities) == 5
    assert result.activity_ids == ("A1790", "A1960", "A2130", "A2280", "A2980")
    assert result.total_amount == pytest.approx(18_936_778.67, abs=0.01)
    expected_first = (
        4_244_054.01 * 0.25
        + 1_522_296.53 * 0.0
        + 5_708_009.02 * 0.50
        + 5_202_860.60 * 0.0
        + 2_259_558.51 * 1.0
    ) / 18_936_778.67
    assert result.period_progress[0] == pytest.approx(expected_first)
    assert result.cumulative_progress[-1] == 1.0


def test_pb2_second_fixed_is_an_independent_exact_name_group():
    service = PaymentBreakdownService()
    rows = (
        _source("A1800", "Second Fixed", 2_829_369.34, (1.0,)),
        _source("A1970", "Second Fixed", 1_014_864.35, (1.0,)),
        _source("A2140", "Second Fixed", 3_805_339.35, (1.0,)),
        _source("A2290", "Second Fixed", 3_468_573.74, (1.0,)),
        _source("A2990", "Second Fixed", 1_506_372.34, (1.0,)),
    )

    result = service.derive_activity(rows)

    assert result.activity_name == "Second Fixed"
    assert len(result.source_activities) == 5
    assert result.total_amount == pytest.approx(12_624_519.12, abs=0.01)


def test_pb2_never_combines_different_names_even_if_they_share_u_glass_keyword():
    service = PaymentBreakdownService()
    rows = (
        _source("A1680", "P3.1-P3.3 | U-Glass", 392_974.86, (1.0,)),
        _source("A2035", "P3.2 U-Glass Wall", 394_550.05, (1.0,)),
    )

    with pytest.raises(PaymentBreakdownError, match="different Activity Names"):
        service.derive_activity(rows)


def test_pb2_group_discovery_defaults_to_repeated_exact_names_only():
    service = PaymentBreakdownService()
    rows = (
        _source("A1", "First Fixed", 60.0, (1.0,)),
        _source("A2", "First Fixed", 40.0, (1.0,)),
        _source("A3", "P3.2 U-Glass Wall", 50.0, (1.0,)),
        _source("A4", "Second Fixed", 70.0, (1.0,)),
        _source("A5", "Second Fixed", 30.0, (1.0,)),
    )

    results = service.derive_exact_name_groups(rows)

    assert [result.activity_name for result in results] == ["First Fixed", "Second Fixed"]


def test_pb2_group_discovery_can_include_singletons_for_diagnostic_review():
    service = PaymentBreakdownService()
    rows = (
        _source("A1", "P3.1-P3.3 | U-Glass", 60.0, (1.0,)),
        _source("A2", "P3.2 U-Glass Wall", 40.0, (1.0,)),
    )

    results = service.derive_exact_name_groups(rows, min_occurrences=1)

    assert [result.activity_name for result in results] == [
        "P3.1-P3.3 | U-Glass",
        "P3.2 U-Glass Wall",
    ]


def test_pb2_case_difference_is_not_exact_name_match():
    service = PaymentBreakdownService()
    rows = (
        _source("A1", "First Fixed", 60.0, (1.0,)),
        _source("A2", "first fixed", 40.0, (1.0,)),
    )

    results = service.derive_exact_name_groups(rows)
    assert results == ()


def test_pb2_rejects_invalid_source_progress_before_aggregation():
    service = PaymentBreakdownService()
    rows = (
        _source("A1", "First Fixed", 60.0, (0.5, 0.25)),
        _source("A2", "First Fixed", 40.0, (1.0,)),
    )

    with pytest.raises(PaymentBreakdownError, match="must total 100%"):
        service.derive_activity(rows)
