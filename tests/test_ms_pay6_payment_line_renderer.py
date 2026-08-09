from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook

from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
from progress_studio.services.payment_service import PaymentService


def _progress_workbook(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "main"
    headers = [
        "Row Type", "WBS", "Description", "P/A", "Activity ID", "Task ID", "UID", "Outline Level",
        "Plan Start", "Plan Finish", "Actual Start", "Actual Finish", "% Complete", "Physical %",
        "Amount", "Total Float (hr)", "XML Amount",
    ]
    for col, value in enumerate(headers, start=1):
        ws.cell(4, col, value)
    week0 = datetime(2026, 3, 2)
    for offset in range(5):
        ws.cell(4, 18 + offset, week0 + timedelta(days=7 * offset))

    ws.cell(5, 1, "Project Summary")
    ws.cell(5, 3, "Demo")
    ws.cell(5, 4, "P")
    ws.cell(5, 9, datetime(2026, 3, 2))
    ws.cell(5, 10, datetime(2026, 4, 3))

    ws.cell(6, 1, "Activity")
    ws.cell(6, 2, "1.1")
    ws.cell(6, 3, "Foundation")
    ws.cell(6, 4, "P")
    ws.cell(6, 5, "A1000")
    ws.cell(6, 9, datetime(2026, 3, 2))
    ws.cell(6, 10, datetime(2026, 4, 3))
    for col, value in zip(range(18, 23), [0.10, 0.15, 0.20, 0.20, 0.35]):
        ws.cell(6, col, value)

    # Keep a spacer/WBS-like row between points so the vertical line must cross rows.
    ws.cell(7, 1, "WBS")
    ws.cell(7, 2, "2")
    ws.cell(7, 3, "Structure")
    ws.cell(7, 4, "P")

    ws.cell(8, 1, "Activity")
    ws.cell(8, 2, "2.1")
    ws.cell(8, 3, "Structure L1")
    ws.cell(8, 4, "P")
    ws.cell(8, 5, "A1010")
    ws.cell(8, 9, datetime(2026, 3, 16))
    ws.cell(8, 10, datetime(2026, 3, 30))
    for col, value in zip(range(20, 23), [0.25, 0.25, 0.50]):
        ws.cell(8, col, value)

    wb.save(path)
    wb.close()
    return path


def _payment_input(progress: Path, path: Path) -> Path:
    PaymentInputWorkbook().create(progress, path, 3)
    wb = load_workbook(path)
    ws = wb["Payment Input"]
    # P01 begins at column E in the lightweight tree layout. Clear generated
    # suggestions first, then make P01 intentionally sparse with two points.
    for row in range(8, ws.max_row + 1):
        if ws.cell(row, 1).value == "ACT":
            for col in range(5, 8):
                ws.cell(row, col).value = None
    activity_rows = {
        ws.cell(row, 3).value: row
        for row in range(8, ws.max_row + 1)
        if ws.cell(row, 1).value == "ACT"
    }
    ws.cell(activity_rows["A1000"], 5, 0.60)
    ws.cell(activity_rows["A1010"], 5, 0.25)
    wb.save(path)
    wb.close()
    return path


def test_ms_pay6_single_period_uses_vertical_backbone_without_shapes(tmp_path: Path) -> None:
    progress = _progress_workbook(tmp_path / "progress.xlsx")
    payment = _payment_input(progress, tmp_path / "payment_input.xlsx")
    output = tmp_path / "payment_p01.xlsx"

    result = PaymentService().render_single_payment_line(progress, payment, output, "P01")

    assert result.period_id == "P01"
    assert result.rendered_points == 2
    assert result.color == "C00000"

    wb = load_workbook(output)
    try:
        assert "Payment" in wb.sheetnames
        main = wb["main"]
        sheet = wb["Payment"]
        assert not sheet._images
        assert not sheet._charts

        # P01 date 13-Mar resolves to the boundary before T (16-Mar bucket).
        # Backbone runs only from first sparse Activity row through the last one.
        assert sheet["T6"].border.left.style == "medium"
        assert sheet["T7"].border.left.style == "medium"
        assert sheet["T8"].border.left.style == "medium"
        assert sheet["T6"].border.left.color.rgb.endswith("C00000")

        # A1000 target = right edge of U => before V; branch T..V on row 6.
        assert sheet["T6"].border.bottom.style == "medium"
        assert sheet["U6"].border.bottom.style == "medium"
        assert sheet["V6"].border.left.style == "medium"

        # A1010 target = right edge of T => before U.
        assert sheet["T8"].border.bottom.style == "medium"
        assert sheet["U8"].border.left.style == "medium"

        # Source of truth is untouched.
        assert main["T6"].border.left.style != "medium"
        assert main["U6"].border.bottom.style != "medium"
    finally:
        wb.close()


def test_ms_pay61_renders_three_colored_vertical_backbones(tmp_path: Path) -> None:
    progress = _progress_workbook(tmp_path / "progress.xlsx")
    payment = _payment_input(progress, tmp_path / "payment_input.xlsx")

    wb = load_workbook(payment)
    ws = wb["Payment Input"]
    activity_rows = {
        ws.cell(row, 3).value: row
        for row in range(8, ws.max_row + 1)
        if ws.cell(row, 1).value == "ACT"
    }
    # Tree layout: E=P01, F=P02, G=P03. Keep all three periods sparse.
    ws.cell(activity_rows["A1000"], 5, 0.25)
    ws.cell(activity_rows["A1000"], 6, 0.60)
    ws.cell(activity_rows["A1000"], 7, 1.00)
    ws.cell(activity_rows["A1010"], 5).value = None
    ws.cell(activity_rows["A1010"], 6, 0.25)
    ws.cell(activity_rows["A1010"], 7, 1.00)
    wb.save(payment)
    wb.close()

    output = tmp_path / "payment_three.xlsx"
    result = PaymentService().render_payment_backbones(
        progress, payment, output, ("P01", "P02", "P03")
    )

    assert result.period_ids == ("P01", "P02", "P03")
    assert result.rendered_periods == 3
    assert result.rendered_points == 5
    assert dict(result.colors) == {"P01": "C00000", "P02": "0070C0", "P03": "548235"}

    wb = load_workbook(output)
    try:
        sheet = wb["Payment"]
        assert not sheet._images
        assert not sheet._charts
        # Backbone payment-date boundaries: P01=T, P02=U, P03=W (after last bucket).
        assert sheet["T6"].border.left.color.rgb.endswith("C00000")
        assert sheet["U6"].border.left.color.rgb.endswith("0070C0")
        # P03's far-right backbone uses right border of V because W is beyond max timeline.
        assert sheet["V6"].border.right.style == "medium"
        assert sheet["V6"].border.right.color.rgb.endswith("548235")
    finally:
        wb.close()
