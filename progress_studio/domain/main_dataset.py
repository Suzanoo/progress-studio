
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any


@dataclass(frozen=True, slots=True)
class MainPeriod:
    column: int
    key: str
    reporting_date: datetime | None


@dataclass(frozen=True, slots=True)
class MainRow:
    row_number: int
    row_type: str
    pa: str
    wbs: str
    description: str
    activity_id: str
    outline_level: int | None
    plan_start: datetime | None
    plan_finish: datetime | None
    amount: float | None
    percent_complete: float | None
    period_values: tuple[tuple[int, float | None], ...]

    def period_value(self, column: int) -> float | None:
        for col, value in self.period_values:
            if col == column:
                return value
        return None


@dataclass(frozen=True, slots=True)
class MainDataset:
    workbook_name: str
    header_row: int
    headers: tuple[tuple[str, int], ...]
    periods: tuple[MainPeriod, ...]
    rows: tuple[MainRow, ...]

    @property
    def activities(self) -> tuple[MainRow, ...]:
        return tuple(
            row for row in self.rows
            if row.row_type.lower() == "activity" and row.pa.upper() == "P" and row.activity_id
        )

    def header_column(self, name: str) -> int | None:
        target = name.strip().lower()
        for header, col in self.headers:
            if header == target:
                return col
        return None
