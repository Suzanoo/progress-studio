
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True, slots=True)
class MonthlyPeriod:
    key: str
    reporting_date: datetime | None
    source_columns: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class MonthlyRow:
    source_row: int
    row_type: str
    pa: str
    wbs: str
    description: str
    activity_id: str
    outline_level: int | None
    values: tuple[float | None, ...]


@dataclass(frozen=True, slots=True)
class MonthlyCache:
    periods: tuple[MonthlyPeriod, ...]
    rows: tuple[MonthlyRow, ...]

    @property
    def value_cell_count(self) -> int:
        return len(self.periods) * len(self.rows)


@dataclass(frozen=True, slots=True)
class MonthlyArchitectureDecision:
    winner: str
    formula_cells: int
    cache_value_cells: int
    direct_render_cells: int
    rationale: tuple[str, ...]
