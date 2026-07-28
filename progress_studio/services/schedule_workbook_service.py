from __future__ import annotations

from pathlib import Path

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.infrastructure.excel import transform_file


class ScheduleWorkbookService:
    def prepare(self, input_file: Path, output_file: Path) -> tuple[int, int, int]:
        return transform_file(input_file, output_file, WORKBOOK_SCHEMA.main_sheet, False)
