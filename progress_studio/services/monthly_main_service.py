from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.infrastructure.excel.monthly_main_workbook import build_monthly_main_view


class MonthlyMainService:
    """Create the calculated monthly counterpart of the weekly main sheet."""

    def build(self, input_file: Path, output_file: Path) -> Path:
        workbook = load_workbook(input_file)
        try:
            build_monthly_main_view(
                workbook,
                source_sheet=WORKBOOK_SCHEMA.main_sheet,
                target_sheet=WORKBOOK_SCHEMA.main_monthly_sheet,
            )
            output_file.parent.mkdir(parents=True, exist_ok=True)
            workbook.save(output_file)
        finally:
            workbook.close()
        return output_file
