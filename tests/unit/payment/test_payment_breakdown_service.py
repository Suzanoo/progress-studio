import pytest

from progress_studio.services.payment_breakdown_service import (
    PaymentBreakdownError,
    PaymentBreakdownItem,
    PaymentBreakdownService,
)


UGLASS_AMOUNT = 3_778_840.27


def _seven_period_profile() -> tuple[float, ...]:
    # 7.14%, 16.67% x 5, 9.52% using exact fractions so the contract totals 100%.
    return (1 / 14, 1 / 6, 1 / 6, 1 / 6, 1 / 6, 1 / 6, 2 / 21)


def test_pb1_u_glass_exact_name_weighted_progress_contract():
    service = PaymentBreakdownService()
    profile = _seven_period_profile()

    # PB-1 golden case: U-Glass with the four payment sub-activities seen in
    # the reference workbook.  Their schedules are deliberately staggered.
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

    # First period contains only item 2, therefore U-Glass progress is the
    # item's progress weighted by item amount / U-Glass amount (~0.746%).
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
