from __future__ import annotations

from pathlib import Path
import shutil

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from progress_studio.domain.mapping_models import (
    ActivityRow as ActivityRecord,
    AllocationRecord,
    BOQRow as BOQRecord,
)
from progress_studio.infrastructure.excel.mapping_reader import BOQSheetReader, ProgressActivityReader


def _headers(ws, row: int = 1) -> dict[str, int]:
    return {
        str(ws.cell(row, col).value or "").strip().lower(): col
        for col in range(1, ws.max_column + 1)
        if str(ws.cell(row, col).value or "").strip()
    }


class BOQMappingService:
    """Read mapping inputs efficiently and export percentage allocations."""

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
        allocations: list[AllocationRecord],
    ) -> Path:
        """Compatibility wrapper; new GUI uses WorkbookExportService."""
        from progress_studio.domain.export_models import ExportValidation
        from progress_studio.infrastructure.excel.mapped_workbook_exporter import MappedWorkbookExporter

        by_key = {row.key: row for row in boq_rows}
        allocated = sum(
            by_key[item.boq_key].amount * item.share_percent / 100.0
            for item in allocations if item.boq_key in by_key
        )
        shares: dict[str, float] = {}
        for item in allocations:
            shares[item.boq_key] = shares.get(item.boq_key, 0.0) + item.share_percent
        total = sum(row.amount for row in boq_rows)
        validation = ExportValidation(
            activity_count=0, boq_count=len(boq_rows), allocation_count=len(allocations),
            mapped_activity_count=len({item.activity_id for item in allocations}),
            mapped_boq_count=sum(value > 1e-9 for value in shares.values()),
            full_boq_count=sum(value >= 100.0 - 1e-9 for value in shares.values()),
            partial_boq_count=sum(1e-9 < value < 100.0 - 1e-9 for value in shares.values()),
            unmapped_boq_count=sum(row.key not in shares for row in boq_rows),
            total_boq_amount=total, allocated_amount=allocated, remaining_amount=max(0.0, total - allocated),
        )
        result = MappedWorkbookExporter().export(
            progress_file, output_file, boq_rows, allocations, validation, overwrite=True
        )
        return result.output_file
