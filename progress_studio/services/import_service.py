from __future__ import annotations

from pathlib import Path

from progress_studio.infrastructure.excel import ImportWorkbookWriter
from progress_studio.infrastructure.schedule_xml import ScheduleXmlReader
from progress_studio.services.schedule_service import ScheduleService


class ImportService:
    def __init__(self, reader: ScheduleXmlReader, schedule_service: ScheduleService, writer: ImportWorkbookWriter) -> None:
        self._reader = reader
        self._schedule_service = schedule_service
        self._writer = writer

    def import_xml(self, source_xml: Path, output_file: Path) -> tuple[str, int, int]:
        project_name, rows = self._reader.read(source_xml)
        self._schedule_service.roll_up_summary_dates(rows)
        self._writer.write(output_file, source_xml, project_name, rows)
        return project_name, sum(row.is_summary for row in rows), sum(not row.is_summary for row in rows)
