from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path


@dataclass(frozen=True)
class PaymentWorkbookValidation:
    workbook: Path
    main_sheet: str
    activity_rows: int
    max_row: int
    max_column: int
    project_start: date | None = None
    project_finish: date | None = None
    default_payment_periods: int = 1


@dataclass(frozen=True)
class PaymentSnapshotResult:
    source_workbook: Path
    output_workbook: Path
    payment_sheet: str
    replaced_existing_sheet: bool
    activity_rows: int


@dataclass(frozen=True)
class PaymentInputResult:
    source_workbook: Path
    output_workbook: Path
    payment_periods: int
    activity_rows: int
    project_start: date
    project_finish: date


@dataclass(frozen=True)
class PaymentInputValidation:
    workbook: Path
    payment_sheet: str
    payment_periods: int
    activity_rows: int
    matched_activities: int
    missing_activities: int
