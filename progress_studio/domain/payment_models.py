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
    populated_requirements: int = 0


@dataclass(frozen=True)
class PaymentRequirement:
    activity_id: str
    required_fraction: float
    source_row: int
    source_column: int


@dataclass(frozen=True)
class PaymentPeriodRequirements:
    period_id: str
    column_index: int
    payment_date: date | None
    requirements: tuple[PaymentRequirement, ...]


@dataclass(frozen=True)
class PaymentInputData:
    workbook: Path
    sheet: str
    periods: tuple[PaymentPeriodRequirements, ...]
    activity_ids: tuple[str, ...]
    populated_requirements: int


@dataclass(frozen=True)
class ActivityProgressBucket:
    column_index: int
    column_letter: str
    week_start: date
    incremental_fraction: float
    cumulative_fraction: float


@dataclass(frozen=True)
class ActivityProgress:
    activity_id: str
    row_number: int
    buckets: tuple[ActivityProgressBucket, ...]


@dataclass(frozen=True)
class ActivityProgressIndex:
    workbook: Path
    sheet: str
    timescale_columns: tuple[tuple[int, str, date], ...]
    activities: dict[str, ActivityProgress]


@dataclass(frozen=True)
class PaymentResolvedPoint:
    period_id: str
    activity_id: str
    required_fraction: float
    activity_row: int
    timescale_column: int
    timescale_column_letter: str
    boundary_edge: str
    week_start: date
    reached_cumulative: float


@dataclass(frozen=True)
class PaymentResolvedPeriod:
    period_id: str
    payment_date: date | None
    points: tuple[PaymentResolvedPoint, ...]


@dataclass(frozen=True)
class PaymentPositionIssue:
    period_id: str
    activity_id: str
    code: str
    message: str


@dataclass(frozen=True)
class PaymentPositionResult:
    periods: tuple[PaymentResolvedPeriod, ...]
    issues: tuple[PaymentPositionIssue, ...]
    requirement_count: int
    resolved_count: int


@dataclass(frozen=True)
class PaymentPreparationResult:
    validation: PaymentInputValidation
    positions: PaymentPositionResult
