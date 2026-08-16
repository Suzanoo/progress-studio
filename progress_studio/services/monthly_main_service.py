from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.infrastructure.excel.final_workbook_policy import finalize_workbook
from progress_studio.infrastructure.excel.calculation_policy import request_initial_manual_excel_recalculation
from progress_studio.infrastructure.excel.main_dataset_workbook_adapter import main_dataset_from_workbook
from progress_studio.infrastructure.excel.monthly_main_workbook import build_monthly_main_view
from progress_studio.infrastructure.excel.traditional_overlay_workbook import build_traditional_overlays


class MonthlyMainService:
    """Final Create Progress boundary: monthly view + proven workbook features."""

    def build(self, input_file: Path, output_file: Path) -> Path:
        # This is already the final openpyxl load in the Create Progress pipeline.
        # Keep the workbook in RAM, render/finalize it here, and save only once.
        workbook = load_workbook(input_file)
        try:
            build_monthly_main_view(
                workbook,
                source_sheet=WORKBOOK_SCHEMA.main_sheet,
                target_sheet=WORKBOOK_SCHEMA.main_monthly_sheet,
            )
            dataset = main_dataset_from_workbook(workbook, workbook_name=output_file.name)
            build_traditional_overlays(workbook, dataset)
            finalize_workbook(workbook, mode="snapshot", include_guide=True)
            # Create Progress needs one Excel-owned full calculation on first open
            # so Dashboard_Data formula caches and charts are complete immediately.
            # The workbook remains Manual afterwards (F9 / Save contract).
            request_initial_manual_excel_recalculation(workbook)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            workbook.save(output_file)
        finally:
            workbook.close()
        return output_file
