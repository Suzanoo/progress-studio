from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta

from progress_studio.domain.main_dataset import MainDataset, MainPeriod


@dataclass(frozen=True, slots=True)
class ReportingPeriodRef:
    """One real reporting period and its original zero-based dataset index."""

    index: int
    period: MainPeriod


def _as_date(value: date | datetime | None) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    return value


def weekly_period_overlaps_project(
    reporting_date: date | datetime | None,
    project_start: date | datetime | None,
    project_finish: date | datetime | None,
) -> bool:
    """Return True when a weekly reporting interval overlaps the project window.

    Progress Studio intentionally displays +/- timescale margin around the real
    schedule.  A reporting cutoff belongs to the real reporting range when its
    seven-day interval intersects Project Start..Project Finish.  This preserves
    the final reporting week when Project Finish falls before that week's cutoff.
    """
    reporting = _as_date(reporting_date)
    start = _as_date(project_start)
    finish = _as_date(project_finish)
    if reporting is None:
        return False
    if start is None or finish is None:
        return True
    period_start = reporting - timedelta(days=6)
    return reporting >= start and period_start <= finish


def project_schedule_window(dataset: MainDataset) -> tuple[date | None, date | None]:
    """Return the editable activity schedule window, excluding display margin."""
    starts = [
        row.plan_start.date()
        for row in dataset.activities
        if row.plan_start is not None
    ]
    finishes = [
        row.plan_finish.date()
        for row in dataset.activities
        if row.plan_finish is not None
    ]
    return (min(starts) if starts else None, max(finishes) if finishes else None)


def select_reporting_periods(dataset: MainDataset) -> tuple[ReportingPeriodRef, ...]:
    """Select project-only weekly periods while preserving source row identity.

    ``dataset.periods`` mirrors the visible ``main`` timescale and therefore may
    contain pre/post margin.  Downstream reporting data (Dashboard_Data, KPI and
    chart sources) must use this selector instead of consuming all display periods.
    """
    project_start, project_finish = project_schedule_window(dataset)
    selected = tuple(
        ReportingPeriodRef(index=index, period=period)
        for index, period in enumerate(dataset.periods)
        if weekly_period_overlaps_project(
            period.reporting_date,
            project_start,
            project_finish,
        )
    )
    # Preserve legacy resilience for malformed/old workbooks whose activity rows
    # do not expose Plan Start/Finish. In that case the selector intentionally
    # falls back to every dated period rather than silently producing no dashboard.
    if selected:
        return selected
    return tuple(
        ReportingPeriodRef(index=index, period=period)
        for index, period in enumerate(dataset.periods)
        if period.reporting_date is not None
    )
