from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class ExportValidation:
    activity_count: int
    boq_count: int
    allocation_count: int
    mapped_activity_count: int
    mapped_boq_count: int
    full_boq_count: int
    partial_boq_count: int
    unmapped_boq_count: int
    total_boq_amount: float
    allocated_amount: float
    remaining_amount: float

    @property
    def is_complete(self) -> bool:
        return self.partial_boq_count == 0 and self.unmapped_boq_count == 0

    @property
    def allocated_percent(self) -> float:
        if self.total_boq_amount <= 0:
            return 0.0
        return self.allocated_amount / self.total_boq_amount * 100.0


@dataclass(frozen=True, slots=True)
class ExportResult:
    output_file: Path
    validation: ExportValidation
    amount_rows_updated: int
    mapping_rows_written: int
