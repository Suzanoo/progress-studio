from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from progress_studio.domain.payment_models import PaymentInputResult, PaymentInputValidation
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError, PaymentWorkbookSnapshotter


class PaymentInputWorkbook:
    """Create/validate the small user-editable Payment Requirement payload."""

    SHEET = "Payment Input"
    HEADER_ROW = 6
    DATE_ROW = 7
    FIRST_ACTIVITY_ROW = 8

    def __init__(self, snapshotter: PaymentWorkbookSnapshotter | None = None) -> None:
        self.snapshotter = snapshotter or PaymentWorkbookSnapshotter()

    def create(self, source: Path, output: Path, periods: int) -> PaymentInputResult:
        source = Path(source)
        output = Path(output)
        if periods < 1 or periods > 120:
            raise PaymentWorkbookError("Payment periods must be between 1 and 120.")

        validation = self.snapshotter.validate(source)
        if validation.project_start is None or validation.project_finish is None:
            raise PaymentWorkbookError("Project Start / Finish could not be read from the main sheet.")

        activity_ids = self.snapshotter.activity_ids(source)
        if not activity_ids:
            raise PaymentWorkbookError("No Activity IDs were found in the main sheet.")
        payment_dates = self._spread_dates(validation.project_start, validation.project_finish, periods)

        wb = Workbook()
        ws = wb.active
        ws.title = self.SHEET
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Payment Requirement Input"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A2"] = "Project Start"
        ws["B2"] = validation.project_start
        ws["A3"] = "Project Finish"
        ws["B3"] = validation.project_finish
        ws["A4"] = "Payment Periods"
        ws["B4"] = periods
        ws["B2"].number_format = ws["B3"].number_format = "dd-mmm-yy"

        ws.cell(self.HEADER_ROW, 1, "Activity ID")
        ws.cell(self.DATE_ROW, 1, "Payment Date")
        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        for idx, payment_date in enumerate(payment_dates, start=1):
            col = idx + 1
            ws.cell(self.HEADER_ROW, col, f"P{idx:02d}")
            ws.cell(self.DATE_ROW, col, payment_date)
            ws.cell(self.DATE_ROW, col).number_format = "dd-mmm-yy"
            ws.column_dimensions[get_column_letter(col)].width = 11
        for cell in ws[self.HEADER_ROW]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        ws.cell(self.DATE_ROW, 1).font = Font(bold=True)

        for row_idx, activity_id in enumerate(activity_ids, start=self.FIRST_ACTIVITY_ROW):
            ws.cell(row_idx, 1, activity_id)
            # Blank means "no requirement". Keep cells percentage-formatted so the
            # user can type 25%, 50%, 100% directly without confusing blank with 0%.
            for col in range(2, periods + 2):
                ws.cell(row_idx, col).number_format = "0%"

        last_col = get_column_letter(periods + 1)
        ws.freeze_panes = "B8"
        ws.auto_filter.ref = f"A{self.HEADER_ROW}:{last_col}{self.FIRST_ACTIVITY_ROW + len(activity_ids) - 1}"
        ws.column_dimensions["A"].width = 18
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
        wb.close()

        return PaymentInputResult(
            source_workbook=source,
            output_workbook=output,
            payment_periods=periods,
            activity_rows=len(activity_ids),
            project_start=validation.project_start,
            project_finish=validation.project_finish,
        )

    def validate(self, payment_workbook: Path, progress_workbook: Path | None = None) -> PaymentInputValidation:
        path = Path(payment_workbook)
        if not path.exists():
            raise PaymentWorkbookError(f"Payment workbook was not found: {path}")
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise PaymentWorkbookError("Select a Payment Requirement .xlsx or .xlsm workbook.")

        wb = load_workbook(path, read_only=True, data_only=False)
        try:
            if self.SHEET not in wb.sheetnames:
                raise PaymentWorkbookError("Worksheet 'Payment Input' was not found.")
            ws = wb[self.SHEET]
            if str(ws.cell(self.HEADER_ROW, 1).value or "").strip().lower() != "activity id":
                raise PaymentWorkbookError("Payment Input header 'Activity ID' was not found.")

            payment_headers: list[str] = []
            col = 2
            while True:
                value = ws.cell(self.HEADER_ROW, col).value
                if value is None:
                    break
                text = str(value).strip().upper()
                if not text.startswith("P"):
                    break
                payment_headers.append(text)
                col += 1
            if not payment_headers:
                raise PaymentWorkbookError("No payment period columns were found.")

            activity_ids: list[str] = []
            row = self.FIRST_ACTIVITY_ROW
            while True:
                value = ws.cell(row, 1).value
                if value is None:
                    break
                text = str(value).strip()
                if text:
                    activity_ids.append(text)
                row += 1
            if not activity_ids:
                raise PaymentWorkbookError("No Activity IDs were found in Payment Input.")
        finally:
            wb.close()

        matched = len(activity_ids)
        missing = 0
        if progress_workbook is not None:
            progress_ids = set(self.snapshotter.activity_ids(Path(progress_workbook)))
            matched = sum(1 for activity_id in activity_ids if activity_id in progress_ids)
            missing = len(activity_ids) - matched

        return PaymentInputValidation(
            workbook=path,
            payment_sheet=self.SHEET,
            payment_periods=len(payment_headers),
            activity_rows=len(activity_ids),
            matched_activities=matched,
            missing_activities=missing,
        )

    @staticmethod
    def _spread_dates(start: date, finish: date, periods: int) -> list[date]:
        total_days = max((finish - start).days, 1)
        return [start + timedelta(days=round(total_days * idx / periods)) for idx in range(1, periods + 1)]
