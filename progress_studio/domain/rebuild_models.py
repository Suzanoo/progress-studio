from dataclasses import dataclass

@dataclass(frozen=True, slots=True)
class EditedWorkbookMigrationResult:
    source_activity_count: int
    target_activity_count: int
    matched_activity_count: int
    matched_by_activity_id: int
    matched_by_signature: int
    unmatched_activity_ids: tuple[str, ...]
    ambiguous_activity_ids: tuple[str, ...]
    amount_cells_migrated: int
    plan_cells_migrated: int
    actual_cells_migrated: int

    @property
    def has_review_items(self) -> bool:
        return bool(self.unmatched_activity_ids or self.ambiguous_activity_ids)

from enum import Enum
from pathlib import Path


class RebuildMode(str, Enum):
    """What the standalone rebuild engine is allowed to regenerate."""

    PROGRESS = "progress"
    PAYMENT = "payment"


@dataclass(frozen=True, slots=True)
class RebuildSheetContract:
    """Immutable sheet ownership contract for standalone workbook rebuild."""

    source_of_truth: tuple[str, ...]
    preserve: tuple[str, ...]
    generated_progress: tuple[str, ...]
    generated_payment: tuple[str, ...]
    internal_preserve: tuple[str, ...]

    def generated_for(self, mode: RebuildMode) -> tuple[str, ...]:
        if mode is RebuildMode.PROGRESS:
            return self.generated_progress
        if mode is RebuildMode.PAYMENT:
            return self.generated_payment
        raise ValueError(f"Unsupported rebuild mode: {mode}")


@dataclass(frozen=True, slots=True)
class RebuildWorkbookAnalysis:
    workbook: Path
    mode: RebuildMode
    main_sheet: str
    main_rows: int
    main_columns: int
    activity_count: int
    payment_input_present: bool
    existing_generated_sheets: tuple[str, ...]
    missing_generated_sheets: tuple[str, ...]
    preserve_sheets_present: tuple[str, ...]
    unknown_sheets: tuple[str, ...]
    contract: RebuildSheetContract

    @property
    def ready(self) -> bool:
        if self.mode is RebuildMode.PAYMENT:
            return self.payment_input_present
        return True

@dataclass(frozen=True, slots=True)
class ProgressRebuildResult:
    source_workbook: Path
    output_workbook: Path
    activity_count: int
    week_count: int
    progress_table_rows: int
    progress_table_checked_cells: int
    monthly_periods: int
    rebuilt_sheets: tuple[str, ...]
    preserved_payment_sheet: bool
    preserved_payment_input_sheet: bool

