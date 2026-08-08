from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PaymentWorkbookValidation:
    workbook: Path
    main_sheet: str
    activity_rows: int
    max_row: int
    max_column: int


@dataclass(frozen=True)
class PaymentSnapshotResult:
    source_workbook: Path
    output_workbook: Path
    payment_sheet: str
    replaced_existing_sheet: bool
    activity_rows: int
