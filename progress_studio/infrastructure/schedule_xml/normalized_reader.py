from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from progress_studio.domain import Activity, NormalizedScheduleValidator

from .format_detector import ScheduleXmlFormat, ScheduleXmlFormatDetector
from .msp_adapter import MspXmlAdapter
from .p6_adapter import P6XmlAdapter


class NormalizedScheduleXmlReader:
    """N-7 production XML boundary for Create Progress.

    Detect the XML dialect, normalize it through the source-specific adapter,
    validate the source-neutral contract, then project it back onto the legacy
    ``Activity`` rows consumed by the proven workbook pipeline.

    The bridge is deliberately temporary and narrow: workbook generation stays
    untouched while XML parsing is removed from the engine boundary.
    """

    def __init__(
        self,
        detector: ScheduleXmlFormatDetector | None = None,
        msp_adapter: MspXmlAdapter | None = None,
        p6_adapter: P6XmlAdapter | None = None,
        validator: NormalizedScheduleValidator | None = None,
    ) -> None:
        self._detector = detector or ScheduleXmlFormatDetector()
        self._msp = msp_adapter or MspXmlAdapter(detector=self._detector)
        self._p6 = p6_adapter or P6XmlAdapter(detector=self._detector)
        self._validator = validator or NormalizedScheduleValidator()

    def read(self, xml_file: Path) -> tuple[str, list[Activity]]:
        xml_file = Path(xml_file)
        detected = self._detector.detect(xml_file)
        if detected is ScheduleXmlFormat.MSP_XML:
            schedule = self._msp.normalize(xml_file)
        elif detected is ScheduleXmlFormat.P6_XML:
            schedule = self._p6.normalize(xml_file)
        else:
            raise ValueError(
                "Unsupported schedule XML format. Progress Studio currently supports "
                "Microsoft Project XML and Primavera P6 XML."
            )

        schedule = self._validator.validate(schedule)
        rows = self._to_legacy_rows(schedule, preserve_source_order=detected is ScheduleXmlFormat.MSP_XML)
        return schedule.project.project_name, rows

    @staticmethod
    def _summary_row(row) -> Activity:
        return Activity(
            source_order=row.source_order,
            task_id=None,
            uid=None,
            activity_id="",
            name=row.wbs_name,
            wbs=row.wbs_code,
            outline_level=row.outline_level,
            is_summary=True,
            plan_start=None,
            plan_finish=None,
            actual_start=None,
            actual_finish=None,
            percent_complete=None,
            physical_percent_complete=None,
            total_slack_minutes=None,
            amount=None,
        )

    @staticmethod
    def _activity_row(row) -> Activity:
        return Activity(
            source_order=row.source_order,
            task_id=None,
            uid=None,
            activity_id=row.activity_id,
            name=row.activity_name,
            wbs=row.wbs_code,
            outline_level=row.outline_level,
            is_summary=False,
            plan_start=row.plan_start,
            plan_finish=row.plan_finish,
            actual_start=row.actual_start,
            actual_finish=row.actual_finish,
            percent_complete=row.percent_complete,
            physical_percent_complete=row.physical_percent_complete,
            total_slack_minutes=None,
            amount=None,
        )

    @classmethod
    def _to_legacy_rows(cls, schedule, *, preserve_source_order: bool) -> list[Activity]:
        if preserve_source_order:
            combined: list[Activity] = [cls._summary_row(row) for row in schedule.wbs]
            combined.extend(cls._activity_row(row) for row in schedule.activities)
            return sorted(combined, key=lambda row: row.source_order)

        # P6 exports WBS and Activity as separate object collections. Rebuild a
        # deterministic hierarchy for the legacy workbook writer: each WBS is
        # emitted before its direct activities/children, and source order is
        # retained within each group.
        wbs_by_code = {row.wbs_code: row for row in schedule.wbs}
        child_wbs: dict[str | None, list] = defaultdict(list)
        for row in schedule.wbs:
            child_wbs[row.parent_wbs_code].append(row)
        for rows in child_wbs.values():
            rows.sort(key=lambda row: row.source_order)

        activities_by_wbs: dict[str, list] = defaultdict(list)
        for row in schedule.activities:
            activities_by_wbs[row.wbs_code].append(row)
        for rows in activities_by_wbs.values():
            rows.sort(key=lambda row: row.source_order)

        output: list[Activity] = []
        next_order = 0

        def emit(code: str) -> None:
            nonlocal next_order
            wbs = wbs_by_code[code]
            summary = cls._summary_row(wbs)
            summary.source_order = next_order
            next_order += 1
            output.append(summary)

            for activity in activities_by_wbs.get(code, ()): 
                legacy = cls._activity_row(activity)
                legacy.source_order = next_order
                next_order += 1
                output.append(legacy)

            for child in child_wbs.get(code, ()): 
                emit(child.wbs_code)

        roots = child_wbs.get(None, [])
        for root in roots:
            emit(root.wbs_code)

        return output
