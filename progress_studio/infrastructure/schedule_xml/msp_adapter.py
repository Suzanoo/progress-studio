from __future__ import annotations

from pathlib import Path

from progress_studio.domain import (
    NormalizedActivity,
    NormalizedProject,
    NormalizedSchedule,
    NormalizedWbs,
)

from .format_detector import ScheduleXmlFormat, ScheduleXmlFormatDetector
from .xml_reader import ScheduleXmlReader


class MspXmlAdapter:
    """Normalize Microsoft Project XML into the source-neutral schedule model.

    N-3 intentionally wraps the proven ``ScheduleXmlReader`` rather than
    replacing its MSP parsing rules.  This keeps Activity ID/Text1 handling,
    task classification, date parsing, and the current MSP import behavior at
    the same boundary while introducing the normalized model for later stages.
    """

    def __init__(
        self,
        reader: ScheduleXmlReader | None = None,
        detector: ScheduleXmlFormatDetector | None = None,
    ) -> None:
        self._reader = reader or ScheduleXmlReader()
        self._detector = detector or ScheduleXmlFormatDetector()

    def normalize(self, xml_file: Path) -> NormalizedSchedule:
        detected = self._detector.detect(xml_file)
        if detected is not ScheduleXmlFormat.MSP_XML:
            raise ValueError(
                f"MspXmlAdapter requires Microsoft Project XML; detected {detected.value}."
            )

        project_name, rows = self._reader.read(xml_file)
        wbs_rows = [row for row in rows if row.is_summary and row.wbs]
        activity_rows = [row for row in rows if not row.is_summary]

        wbs = tuple(
            NormalizedWbs(
                source_order=row.source_order,
                wbs_code=row.wbs,
                wbs_name=row.name,
                parent_wbs_code=self._parent_wbs_code(row.wbs),
                outline_level=row.outline_level,
            )
            for row in wbs_rows
        )

        activities = tuple(
            NormalizedActivity(
                source_order=row.source_order,
                activity_id=row.activity_id,
                activity_name=row.name,
                wbs_code=row.wbs,
                outline_level=row.outline_level,
                plan_start=row.plan_start,
                plan_finish=row.plan_finish,
                actual_start=row.actual_start,
                actual_finish=row.actual_finish,
                percent_complete=row.percent_complete,
                physical_percent_complete=row.physical_percent_complete,
            )
            for row in activity_rows
        )

        plan_starts = [row.plan_start for row in activity_rows if row.plan_start is not None]
        plan_finishes = [row.plan_finish for row in activity_rows if row.plan_finish is not None]
        project = NormalizedProject(
            project_id=None,
            project_name=project_name,
            plan_start=min(plan_starts) if plan_starts else None,
            plan_finish=max(plan_finishes) if plan_finishes else None,
        )

        return NormalizedSchedule(project=project, wbs=wbs, activities=activities)

    @staticmethod
    def _parent_wbs_code(wbs_code: str) -> str | None:
        code = wbs_code.strip()
        if not code or "." not in code:
            return None
        parent = code.rsplit(".", 1)[0].strip()
        return parent or None
