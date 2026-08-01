from __future__ import annotations

from typing import Protocol

from progress_studio.domain.activity import Activity


class ScheduleSource(Protocol):
    """Provide one normalized schedule model to the workbook generator."""

    @property
    def project_name(self) -> str:
        ...

    def activities(self) -> list[Activity]:
        ...
