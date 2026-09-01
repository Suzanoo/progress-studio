from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True, slots=True)
class EarnedValuePoint:
    period_key: str
    reporting_date: datetime | None
    planned_value: float | None
    earned_value: float | None
    schedule_variance: float | None
    schedule_performance_index: float | None


@dataclass(frozen=True, slots=True)
class ActivityEarnedValue:
    activity_id: str
    description: str
    wbs: str
    bac: float
    points: tuple[EarnedValuePoint, ...]


@dataclass(frozen=True, slots=True)
class BOQEarnedValue:
    boq_key: str
    stable_id: str
    description: str
    bac: float
    points: tuple[EarnedValuePoint, ...]


@dataclass(frozen=True, slots=True)
class EarnedValueResult:
    cutoff_date: datetime | None
    project_bac: float
    project_points: tuple[EarnedValuePoint, ...]
    activities: tuple[ActivityEarnedValue, ...]
    boq_items: tuple[BOQEarnedValue, ...]
