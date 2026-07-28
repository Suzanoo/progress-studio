from __future__ import annotations

from pathlib import Path

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.domain.amount import AmountApplicationResult, AmountMappingResult
from progress_studio.infrastructure.excel.amount_workbook import apply_amount_mapping, rebuild_amount_mapping


class AmountService:
    def build_mapping(self, workbook: Path, placeholder: float) -> AmountMappingResult:
        return rebuild_amount_mapping(
            workbook,
            main_sheet=WORKBOOK_SCHEMA.main_sheet,
            mapping_sheet=WORKBOOK_SCHEMA.mapping_sheet,
            placeholder=placeholder,
        )

    def apply_mapping(self, input_file: Path, output_file: Path) -> AmountApplicationResult:
        return apply_amount_mapping(
            input_file,
            output_file,
            main_sheet=WORKBOOK_SCHEMA.main_sheet,
            mapping_sheet=WORKBOOK_SCHEMA.mapping_sheet,
        )
