from pathlib import Path

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.domain.okd import OkdExportResult
from progress_studio.infrastructure.excel.okd_workbook import build_okd_sheets


class OkdService:
    def build(self, input_file: Path, output_file: Path) -> OkdExportResult:
        activities, weeks, table_rows, checked_links = build_okd_sheets(
            input_file,
            output_file,
            source_sheet=WORKBOOK_SCHEMA.main_sheet,
        )
        return OkdExportResult(
            activities=activities,
            weeks=weeks,
            table_rows=table_rows,
            checked_links=checked_links,
        )
