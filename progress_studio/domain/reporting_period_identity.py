from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Sequence

DateLike = date | datetime


def _as_date(value: DateLike) -> date:
    if isinstance(value, datetime):
        return value.date()
    return value


@dataclass(frozen=True, slots=True)
class ReportingPeriodWindow:
    """A physical period window independent of any Excel label or column.

    The window is inclusive on both ends.  It deliberately contains no W/M
    number because reporting labels are presentation metadata, not calculation
    identity.
    """

    start: date
    end: date

    def __post_init__(self) -> None:
        if self.end < self.start:
            raise ValueError("reporting period end must be on or after start")

    @classmethod
    def from_values(cls, start: DateLike, end: DateLike) -> "ReportingPeriodWindow":
        return cls(start=_as_date(start), end=_as_date(end))

    def overlaps(self, project_start: date, project_finish: date) -> bool:
        return self.end >= project_start and self.start <= project_finish


@dataclass(frozen=True, slots=True)
class ReportingPeriodIdentity:
    """Sequential human-facing identity for one physical display period.

    Margin periods intentionally have ``sequence`` and ``label`` set to None.
    Engines should use dates/columns for calculation; W/M labels exist only for
    workbook presentation and user orientation.
    """

    source_index: int
    window: ReportingPeriodWindow
    sequence: int | None
    label: str | None

    @property
    def is_reporting(self) -> bool:
        return self.sequence is not None


def number_reporting_periods(
    periods: Sequence[ReportingPeriodWindow],
    project_start: DateLike,
    project_finish: DateLike,
    *,
    prefix: str,
) -> tuple[ReportingPeriodIdentity, ...]:
    """Assign contiguous labels only to periods overlapping the project window.

    This is a pure numbering utility.  It does not mutate ``periods``, perform
    workbook I/O, infer project boundaries, or decide where physical timescale
    columns live.  Callers own those concerns.
    """

    start = _as_date(project_start)
    finish = _as_date(project_finish)
    if finish < start:
        raise ValueError("project_finish must be on or after project_start")

    clean_prefix = prefix.strip()
    if not clean_prefix:
        raise ValueError("prefix must not be blank")

    identities: list[ReportingPeriodIdentity] = []
    sequence = 0
    for source_index, period in enumerate(periods):
        if period.overlaps(start, finish):
            sequence += 1
            identities.append(
                ReportingPeriodIdentity(
                    source_index=source_index,
                    window=period,
                    sequence=sequence,
                    label=f"{clean_prefix}{sequence}",
                )
            )
        else:
            identities.append(
                ReportingPeriodIdentity(
                    source_index=source_index,
                    window=period,
                    sequence=None,
                    label=None,
                )
            )

    return tuple(identities)
