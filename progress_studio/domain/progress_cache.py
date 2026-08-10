
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True, slots=True)
class ProgressCachePoint:
    period_key: str
    reporting_date: datetime | None
    plan_weekly: float | None
    plan_cumulative: float | None
    actual_weekly: float | None
    actual_cumulative: float | None


@dataclass(frozen=True, slots=True)
class ProgressCache:
    total_amount: float
    points: tuple[ProgressCachePoint, ...]

    @property
    def period_count(self) -> int:
        return len(self.points)
