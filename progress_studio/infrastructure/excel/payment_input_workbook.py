from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from progress_studio.domain.payment_models import PaymentInputResult, PaymentInputValidation
from progress_studio.infrastructure.excel.payment_input_reader import PaymentInputSparseReader
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookError, PaymentWorkbookSnapshotter


class PaymentInputWorkbook:
    """Create/validate the small user-editable Payment Requirement payload."""

    SHEET = "Payment Input"
    HEADER_ROW = 6
    DATE_ROW = 7
    FIRST_ACTIVITY_ROW = 8

    def __init__(
        self,
        snapshotter: PaymentWorkbookSnapshotter | None = None,
        reader: PaymentInputSparseReader | None = None,
    ) -> None:
        self.snapshotter = snapshotter or PaymentWorkbookSnapshotter()
        self.reader = reader or PaymentInputSparseReader()

    def create(self, source: Path, output: Path, periods: int) -> PaymentInputResult:
        source = Path(source)
        output = Path(output)
        if periods < 1 or periods > 120:
            raise PaymentWorkbookError("Payment periods must be between 1 and 120.")

        validation = self.snapshotter.validate(source)
        if validation.project_start is None or validation.project_finish is None:
            raise PaymentWorkbookError("Project Start / Finish could not be read from the main sheet.")

        tree_rows = self.snapshotter.payment_tree_rows(source)
        activity_rows = [row for row in tree_rows if row["row_type"] == "ACT"]
        if not activity_rows:
            raise PaymentWorkbookError("No Activity rows were found in the main sheet.")
        payment_dates = self._spread_dates(validation.project_start, validation.project_finish, periods)

        wb = Workbook()
        ws = wb.active
        ws.title = self.SHEET
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.outlinePr.summaryBelow = False

        ws["A1"] = "Payment Requirement Input"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A2"] = "Project Start"
        ws["B2"] = validation.project_start
        ws["A3"] = "Project Finish"
        ws["B3"] = validation.project_finish
        ws["A4"] = "Payment Periods"
        ws["B4"] = periods
        ws["B2"].number_format = ws["B3"].number_format = "dd-mmm-yy"

        fixed_headers = ("Type", "WBS", "Activity ID", "Activity Name")
        for col, label in enumerate(fixed_headers, start=1):
            ws.cell(self.HEADER_ROW, col, label)
        ws.cell(self.DATE_ROW, 1, "Payment Date")

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        for idx, payment_date in enumerate(payment_dates, start=1):
            col = idx + len(fixed_headers)
            ws.cell(self.HEADER_ROW, col, f"P{idx:02d}")
            ws.cell(self.DATE_ROW, col, payment_date)
            ws.cell(self.DATE_ROW, col).number_format = "dd-mmm-yy"
            ws.column_dimensions[get_column_letter(col)].width = 11
        for cell in ws[self.HEADER_ROW]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        ws.cell(self.DATE_ROW, 1).font = Font(bold=True)

        wbs_fills = {
            1: PatternFill("solid", fgColor="F4B183"),
            2: PatternFill("solid", fgColor="F8CBAD"),
            3: PatternFill("solid", fgColor="FCE4D6"),
            4: PatternFill("solid", fgColor="FFF2CC"),
        }
        default_wbs_fill = PatternFill("solid", fgColor="F2F2F2")

        out_row = self.FIRST_ACTIVITY_ROW
        for item in tree_rows:
            level = max(int(item.get("outline_level") or 0), 0)
            ws.row_dimensions[out_row].outlineLevel = min(level, 7)
            if item["row_type"] == "WBS":
                ws.cell(out_row, 1, "WBS")
                ws.cell(out_row, 2, item["wbs"])
                ws.cell(out_row, 4, item["activity_name"])
                fill = wbs_fills.get(level, default_wbs_fill)
                for col in range(1, periods + len(fixed_headers) + 1):
                    cell = ws.cell(out_row, col)
                    cell.fill = fill
                    cell.font = Font(bold=True)
                ws.cell(out_row, 4).alignment = Alignment(indent=max(level - 1, 0))
            else:
                ws.cell(out_row, 1, "ACT")
                ws.cell(out_row, 2, item["wbs"])
                ws.cell(out_row, 3, item["activity_id"])
                ws.cell(out_row, 4, item["activity_name"])
                ws.cell(out_row, 4).alignment = Alignment(indent=max(level - 1, 0))
                fake_values = self._fake_requirements(item["weekly_plan"], payment_dates, validation.project_start)
                for idx, value in enumerate(fake_values, start=1):
                    cell = ws.cell(out_row, idx + len(fixed_headers))
                    cell.number_format = "0%"
                    if value is not None:
                        cell.value = value
            out_row += 1

        ws.freeze_panes = "E8"
        # Intentionally no AutoFilter: hierarchy rows should stay visually attached to their children.
        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 42
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
        wb.close()

        return PaymentInputResult(
            source_workbook=source,
            output_workbook=output,
            payment_periods=periods,
            activity_rows=len(activity_rows),
            project_start=validation.project_start,
            project_finish=validation.project_finish,
        )

    def validate(self, payment_workbook: Path, progress_workbook: Path | None = None) -> PaymentInputValidation:
        data = self.reader.read(Path(payment_workbook))

        matched = len(data.activity_ids)
        missing = 0
        if progress_workbook is not None:
            progress_ids = set(self.snapshotter.activity_ids(Path(progress_workbook)))
            matched = sum(1 for activity_id in data.activity_ids if activity_id in progress_ids)
            missing = len(data.activity_ids) - matched

        return PaymentInputValidation(
            workbook=data.workbook,
            payment_sheet=data.sheet,
            payment_periods=len(data.periods),
            activity_rows=len(data.activity_ids),
            matched_activities=matched,
            missing_activities=missing,
            populated_requirements=data.populated_requirements,
        )

    @staticmethod
    def _fake_requirements(weekly_plan, payment_dates: list[date], project_start: date) -> list[float | None]:
        """Suggested cumulative requirements, but only for periods with planned movement.

        A period gets a value only when the activity has incremental Plan progress
        after the previous payment cut and on/before the current cut. This keeps
        the generated workbook sparse while ensuring short activities are not lost.
        """
        points = sorted((d, float(v)) for d, v in weekly_plan if v is not None)
        result: list[float | None] = []
        cumulative = 0.0
        previous = project_start - timedelta(days=1)
        point_index = 0
        for cut in payment_dates:
            moved = False
            while point_index < len(points) and points[point_index][0] <= cut:
                d, value = points[point_index]
                cumulative += value
                if d > previous and value != 0:
                    moved = True
                point_index += 1
            result.append(min(max(cumulative, 0.0), 1.0) if moved else None)
            previous = cut
        return result

    @staticmethod
    def _spread_dates(start: date, finish: date, periods: int) -> list[date]:
        total_days = max((finish - start).days, 1)
        return [start + timedelta(days=round(total_days * idx / periods)) for idx in range(1, periods + 1)]
