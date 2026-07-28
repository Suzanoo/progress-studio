from pathlib import Path

from openpyxl import load_workbook

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.domain.progress import ProgressBuildResult
from progress_studio.infrastructure.excel.progress_workbook import (
    find_sheet,
    prepare_progress_and_scurve,
)


class ProgressService:
    def build(self, input_file: Path, output_file: Path) -> ProgressBuildResult:
        wb = load_workbook(input_file)
        try:
            ws = find_sheet(wb, WORKBOOK_SCHEMA.main_sheet)
            values = prepare_progress_and_scurve(wb, ws)
            wb.calculation.calcMode = "auto"
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            output_file.parent.mkdir(parents=True, exist_ok=True)
            wb.save(output_file)
        finally:
            wb.close()
        return ProgressBuildResult(*values)
