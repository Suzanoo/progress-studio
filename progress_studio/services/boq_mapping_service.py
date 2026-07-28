from __future__ import annotations

from pathlib import Path
import shutil

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from progress_studio.domain.mapping_models import ActivityRow as ActivityRecord
from progress_studio.domain.mapping_models import BOQRow as BOQRecord
from progress_studio.infrastructure.excel.mapping_reader import BOQSheetReader, ProgressActivityReader


def _headers(ws, row: int = 1) -> dict[str, int]:
    return {
        str(ws.cell(row, col).value or "").strip().lower(): col
        for col in range(1, ws.max_column + 1)
        if str(ws.cell(row, col).value or "").strip()
    }


class BOQMappingService:
    """Read mapping inputs efficiently and export a safe one-item-to-one-activity mapping."""

    def __init__(
        self,
        activity_reader: ProgressActivityReader | None = None,
        boq_reader: BOQSheetReader | None = None,
    ) -> None:
        self.activity_reader = activity_reader or ProgressActivityReader()
        self.boq_reader = boq_reader or BOQSheetReader()

    def read_activities(self, progress_file: Path) -> list[ActivityRecord]:
        return self.activity_reader.read(progress_file)

    def list_boq_sheets(self, boq_file: Path) -> list[str]:
        return self.boq_reader.list_sheets(boq_file)

    def read_boq(self, boq_file: Path, sheet_name: str) -> list[BOQRecord]:
        return self.boq_reader.read(boq_file, sheet_name)

    def export(
        self,
        progress_file: Path,
        output_file: Path,
        boq_rows: list[BOQRecord],
        assignments: dict[str, str],
    ) -> Path:
        output_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(progress_file, output_file)
        wb = load_workbook(output_file)
        try:
            ws_amount = wb["Amount Mapping"]
            h = _headers(ws_amount)
            aid_col = h["activity id"]
            amount_col = h["amount"]
            status_col = h.get("status")

            totals: dict[str, float] = {}
            by_key = {row.key: row for row in boq_rows}
            for key, activity_id in assignments.items():
                row = by_key.get(key)
                if row:
                    totals[activity_id] = totals.get(activity_id, 0.0) + row.amount

            for r in range(2, ws_amount.max_row + 1):
                aid = str(ws_amount.cell(r, aid_col).value or "").strip()
                if not aid:
                    continue
                ws_amount.cell(r, amount_col).value = totals.get(aid, 0.0)
                ws_amount.cell(r, amount_col).number_format = "#,##0.00"
                if status_col:
                    ws_amount.cell(r, status_col).value = "MAPPED" if aid in totals else "UNMAPPED"

            if "BOQ Activity Mapping" in wb.sheetnames:
                del wb["BOQ Activity Mapping"]
            ws = wb.create_sheet("BOQ Activity Mapping")
            headers = [
                "Activity ID", "BOQ Key", "Source Sheet", "Source Row",
                "WBS-2", "WBS-3", "WBS-4", "BOQ Description", "Amount",
            ]
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4472C4")
                cell.alignment = Alignment(horizontal="center")
            for row in boq_rows:
                aid = assignments.get(row.key)
                if not aid:
                    continue
                ws.append([
                    aid, row.key, row.source_sheet, row.source_row,
                    row.wbs2, row.wbs3, row.wbs4, row.description, row.amount,
                ])
                ws.cell(ws.max_row, 9).number_format = "#,##0.00"
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:I{max(1, ws.max_row)}"
            widths = {"A": 16, "B": 30, "C": 22, "D": 12, "E": 22, "F": 28, "G": 28, "H": 60, "I": 18}
            for col, width in widths.items():
                ws.column_dimensions[col].width = width
            wb.calculation.calcMode = "auto"
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.save(output_file)
            return output_file
        finally:
            wb.close()
