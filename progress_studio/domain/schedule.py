from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class ScheduleWindow:
    plan_start: datetime | None
    plan_finish: datetime | None
    actual_start: datetime | None
    actual_finish: datetime | None
