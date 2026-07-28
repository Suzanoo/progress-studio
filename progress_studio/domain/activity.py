from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass
class Activity:
    source_order: int
    task_id: int | None
    uid: int | None
    activity_id: str
    name: str
    wbs: str
    outline_level: int
    is_summary: bool
    plan_start: datetime | None
    plan_finish: datetime | None
    actual_start: datetime | None
    actual_finish: datetime | None
    percent_complete: float | None
    physical_percent_complete: float | None
    total_slack_minutes: float | None
    amount: float | None
