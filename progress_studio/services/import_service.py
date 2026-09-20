from __future__ import annotations

from pathlib import Path

from progress_studio.infrastructure.excel import ImportWorkbookWriter
from progress_studio.infrastructure.schedule_xml import NormalizedScheduleXmlReader
from progress_studio.services.schedule_service import ScheduleService
from progress_studio.services.xml_amount_service import assign_xml_amounts
from progress_studio.services.weighting import assign_dummy_weights, validate_weight_basis


class ImportService:
    def __init__(self, reader: NormalizedScheduleXmlReader, schedule_service: ScheduleService, writer: ImportWorkbookWriter) -> None:
        self._reader = reader
        self._schedule_service = schedule_service
        self._writer = writer

    def import_xml(self, source_xml: Path, output_file: Path, *, weight_basis: str | None = None, amount_field: str | None = None) -> tuple[str, int, int]:
        if weight_basis is not None:
            weight_basis = validate_weight_basis(weight_basis)
        selected = None
        if weight_basis == "amount":
            project_name, rows, fields = self._reader.read_with_amount_fields(source_xml)
            selected = assign_xml_amounts(rows, fields, amount_field).field
        else:
            project_name, rows = self._reader.read(source_xml)
            if weight_basis is not None:
                assign_dummy_weights(rows, weight_basis)
        self._schedule_service.roll_up_summary_dates(rows)
        if weight_basis is None:
            self._writer.write(output_file, source_xml, project_name, rows)
        elif selected is not None:
            self._writer.write(output_file, source_xml, project_name, rows,
                               weight_basis=weight_basis, amount_field=selected)
        else:
            self._writer.write(output_file, source_xml, project_name, rows, weight_basis=weight_basis)
        return project_name, sum(row.is_summary for row in rows), sum(not row.is_summary for row in rows)
