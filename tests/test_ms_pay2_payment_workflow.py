from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from progress_studio.infrastructure.excel.payment_input_workbook import PaymentInputWorkbook
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookSnapshotter


def _progress_workbook(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "main"
    headers = ["Row Type", "WBS", "Description", "P/A", "Activity ID", "Task ID", "UID", "Outline Level", "Plan Start", "Plan Finish"]
    for col, value in enumerate(headers, start=1):
        ws.cell(4, col, value)
    ws.append(["Project Summary", None, "Demo", "P", None, None, None, 0, datetime(2026, 2, 23), datetime(2027, 5, 31)])
    ws.append(["Activity", "1.1", "Mobilization", "P", "A1000", 1, 1, 1, datetime(2026, 2, 23), datetime(2026, 4, 23)])
    ws.append(["Activity", "1.2", "Foundation", "P", "A1010", 2, 2, 1, datetime(2026, 4, 1), datetime(2026, 8, 1)])
    ws.append(["Activity", "1.3", "Structure", "P", "A1020", 3, 3, 1, datetime(2026, 7, 1), datetime(2027, 2, 1)])
    wb.save(path)
    wb.close()
    return path


def test_progress_validation_calculates_default_periods_from_project_dates(tmp_path: Path) -> None:
    source = _progress_workbook(tmp_path / "progress.xlsx")
    result = PaymentWorkbookSnapshotter().validate(source)
    assert result.project_start.isoformat() == "2026-02-23"
    assert result.project_finish.isoformat() == "2027-05-31"
    assert result.default_payment_periods == 15


def test_generate_fake_payment_input_is_single_sheet_and_lightweight(tmp_path: Path) -> None:
    source = _progress_workbook(tmp_path / "progress.xlsx")
    output = tmp_path / "payment_input.xlsx"
    result = PaymentInputWorkbook().create(source, output, 15)

    assert result.payment_periods == 15
    assert result.activity_rows == 3

    wb = load_workbook(output, data_only=False)
    try:
        assert wb.sheetnames == ["Payment Input"]
        ws = wb["Payment Input"]
        assert ws["A6"].value == "Activity ID"
        assert ws["B6"].value == "P01"
        assert ws["P6"].value == "P15"
        assert ws["A8"].value == "A1000"
        assert ws["A10"].value == "A1020"
        assert ws["P7"].value is not None
        assert ws["P8"].value is None
        assert ws["P8"].number_format == "0%"
    finally:
        wb.close()


def test_validate_payment_input_matches_progress_activity_ids(tmp_path: Path) -> None:
    source = _progress_workbook(tmp_path / "progress.xlsx")
    payment = tmp_path / "payment_input.xlsx"
    PaymentInputWorkbook().create(source, payment, 15)

    result = PaymentInputWorkbook().validate(payment, source)
    assert result.payment_periods == 15
    assert result.activity_rows == 3
    assert result.matched_activities == 3
    assert result.missing_activities == 0


def test_validate_payment_input_reports_missing_activity(tmp_path: Path) -> None:
    source = _progress_workbook(tmp_path / "progress.xlsx")
    payment = tmp_path / "payment_input.xlsx"
    PaymentInputWorkbook().create(source, payment, 15)

    wb = load_workbook(payment)
    ws = wb["Payment Input"]
    ws["A9"] = "A9999"
    wb.save(payment)
    wb.close()

    result = PaymentInputWorkbook().validate(payment, source)
    assert result.matched_activities == 2
    assert result.missing_activities == 1
