from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class NormalizedProject:
    """Source-neutral project identity used by the schedule import boundary."""

    project_id: str | None
    project_name: str
    plan_start: datetime | None = None
    plan_finish: datetime | None = None


@dataclass(frozen=True)
class NormalizedWbs:
    """One canonical WBS node.

    ``wbs_code`` deliberately excludes a source-specific project prefix.  For
    example, P6 ``007.2.1.1`` and MSP ``2.1.1`` normalize to ``2.1.1``.
    """

    source_order: int
    wbs_code: str
    wbs_name: str
    parent_wbs_code: str | None
    outline_level: int


@dataclass(frozen=True)
class NormalizedActivity:
    """One source-neutral schedule activity.

    Real Amount is intentionally absent in MS-1. Adapters normalize working
    duration to hours and preserve explicit milestone identity; Create assigns
    Equal/Duration calculation weights separately from schedule normalization.
    """

    source_order: int
    activity_id: str
    activity_name: str
    wbs_code: str
    outline_level: int
    plan_start: datetime | None
    plan_finish: datetime | None
    actual_start: datetime | None = None
    actual_finish: datetime | None = None
    percent_complete: float | None = None
    physical_percent_complete: float | None = None
    duration_hours: float | None = None
    is_milestone: bool = False


@dataclass(frozen=True)
class NormalizedSchedule:
    """Canonical schedule contract consumed by later Progress Studio stages."""

    project: NormalizedProject
    wbs: tuple[NormalizedWbs, ...]
    activities: tuple[NormalizedActivity, ...]
