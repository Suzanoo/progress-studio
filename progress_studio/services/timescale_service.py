from __future__ import annotations

from pathlib import Path

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.infrastructure.excel.timescale_workbook import (
    add_weekly_timescale,
    parse_cutoff_day,
)


class TimescaleService:
    def __init__(self, margin_weeks: int = 4) -> None:
        self.margin_weeks = margin_weeks

    def build(self, input_file: Path, output_file: Path, cutoff_day: str | int) -> Path:
        add_weekly_timescale(
            input_file=input_file,
            output_file=output_file,
            sheet_name=WORKBOOK_SCHEMA.main_sheet,
            cutoff_day=parse_cutoff_day(cutoff_day),
            margin_weeks=self.margin_weeks,
        )
        return output_file
