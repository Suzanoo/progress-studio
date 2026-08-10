
from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(frozen=True, slots=True)
class ActivityTableRow:
    row_type: str
    wbs: str
    activity: str
    activity_id: str
    type_label: str
    total: float | None
    amount: float | None
    progress: float
    variance: float | None
    status: str
    outline_level: int
    source_plan_row: int
    source_actual_row: int | None


@dataclass(frozen=True, slots=True)
class ActivityTableModel:
    cutoff: date | None
    rows: tuple[ActivityTableRow, ...]

    @property
    def pair_count(self) -> int:
        return len(self.rows) // 2
