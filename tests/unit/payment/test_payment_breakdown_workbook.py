from datetime import datetime

import pytest
from openpyxl import Workbook

from progress_studio.domain.main_dataset import MainPeriod
from progress_studio.infrastructure.excel.payment_breakdown_workbook import (
    PAYMENT_BREAKDOWN_SHEET,
    render_payment_breakdown,
)
from progress_studio.services.payment_breakdown_adapter import (
    PaymentBreakdownDatasetSnapshot,
)
from progress_studio.services.payment_breakdown_service import (
    PaymentBreakdownService,
    PaymentBreakdownSourceActivity,
)


def _derived():
    return PaymentBreakdownService().derive_activity(
        (
            PaymentBreakdownSourceActivity(
                "A1",
                "First Fixed",
                100.0,
                (0.0, 0.5, 0.5),
                "3.2",
            ),
            PaymentBreakdownSourceActivity(
                "A2",
                "First Fixed",
                300.0,
                (0.0, 0.0, 1.0),
                "3.3",
            ),
        )
    )


def _snapshot():
    periods = (
        MainPeriod(6, "W1", datetime(2026, 1, 2)),
        MainPeriod(7, "W2", datetime(2026, 1, 9)),
        MainPeriod(8, "W3", datetime(2026, 1, 16)),
    )
    return PaymentBreakdownDatasetSnapshot(
        periods=periods,
        activities=(_derived(),),
        eligible_source_count=2,
        skipped_activity_ids=(),
    )


def _rgb(cell):
    color = cell.font.color
    return None if color is None else color.rgb


def test_pb3_renderer_replaces_only_target_sheet_and_keeps_block_structure():
    wb = Workbook()
    main = wb.active
    main.title = "main"
    old = wb.create_sheet(PAYMENT_BREAKDOWN_SHEET)
    old["A1"] = "old"

    ws = render_payment_breakdown(wb, _snapshot())

    assert wb.sheetnames == ["main", PAYMENT_BREAKDOWN_SHEET]
    assert ws["A1"].value == "Payment Breakdown"
    assert ws["A4"].value == "Activity Name"
    assert ws["A5"].value == "First Fixed"
    assert ws["E5"].value == "Activity Progress"
    assert ws["E6"].value == "Activity Cumulative"
    assert ws["E9"].value == "Combined Progress"
    assert ws["E10"].value == "Combined Cumulative"


def test_pb3_renderer_marks_only_in_progress_percentages_red():
    wb = Workbook()
    wb.active.title = "main"
    ws = render_payment_breakdown(wb, _snapshot())

    # A1 progress row: 0%, 50%, 50%
    assert not (_rgb(ws["F5"]) or "").endswith("FF0000")
    assert (_rgb(ws["G5"]) or "").endswith("FF0000")
    assert (_rgb(ws["H5"]) or "").endswith("FF0000")

    # A2 cumulative row reaches exactly 100% at W3: must not be red.
    assert not (_rgb(ws["H8"]) or "").endswith("FF0000")

    # Combined cumulative is 12.5% at W2 then 100% at W3.
    assert (_rgb(ws["G10"]) or "").endswith("FF0000")
    assert not (_rgb(ws["H10"]) or "").endswith("FF0000")


def test_pb3_renderer_writes_amount_weighted_combined_progress_snapshot():
    wb = Workbook()
    wb.active.title = "main"
    ws = render_payment_breakdown(wb, _snapshot())

    assert ws["D9"].value == pytest.approx(400.0)
    assert ws["F9"].value == pytest.approx(0.0)
    assert ws["G9"].value == pytest.approx(0.125)
    assert ws["H9"].value == pytest.approx(0.875)
    assert ws["H10"].value == pytest.approx(1.0)
