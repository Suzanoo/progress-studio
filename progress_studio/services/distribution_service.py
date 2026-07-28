from pathlib import Path

from progress_studio.config import WORKBOOK_SCHEMA
from progress_studio.domain.distribution import DistributionResult
from progress_studio.infrastructure.excel.distribution_workbook import generate_plan_distribution


class DistributionService:
    def generate(
        self,
        input_file: Path,
        output_file: Path,
        *,
        method: str,
        rules_file: Path | None = None,
        debug: bool = False,
    ) -> DistributionResult:
        generated, no_dates, outside, counts = generate_plan_distribution(
            input_file,
            output_file,
            method=method,
            rules_file=rules_file,
            sheet_name=WORKBOOK_SCHEMA.main_sheet,
            header_row=4,
            week_row=3,
            debug=debug,
        )
        return DistributionResult(method, generated, no_dates, outside, counts)
