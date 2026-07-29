from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

from progress_studio.domain import Activity
from progress_studio.domain.activity_id import ensure_activity_id

from .errors import ScheduleXmlIssue, ScheduleXmlValidationError

P6_ACTIVITY_ID_FIELD_IDS = {"188743731"}
_TASK_TAGS = {"Task", "Activity"}
_NAME_TAGS = ("ActivityName", "TaskName", "Name", "Description")
_START_TAGS = ("PlannedStart", "PlanStart", "Start", "StartDate")
_FINISH_TAGS = ("PlannedFinish", "PlanFinish", "Finish", "FinishDate")
_ACTIVITY_ID_TAGS = (
    "ActivityID",
    "ActivityId",
    "ActivityCode",
    "ActivityCodeValue",
    "TaskID",
    "TaskId",
)
_WBS_TAGS = ("WBS", "WBSCode", "OutlineNumber")


class ScheduleXmlReader:
    """Read schedule XML into the application's existing Activity model.

    The reader is source-tolerant but contract-strict. XML tag aliases may vary,
    while every leaf activity must resolve Activity Name, Plan Start, and Plan Finish.
    """

    def read(self, xml_file: Path) -> tuple[str, list[Activity]]:
        try:
            root = ET.parse(xml_file).getroot()
        except (ET.ParseError, OSError) as exc:
            raise ScheduleXmlValidationError(
                [ScheduleXmlIssue(0, xml_file.name, "Activities", f"Invalid XML file: {exc}")]
            ) from exc

        project_name = self._direct_text(root, ("Name", "ProjectName", "Title")) or xml_file.stem
        nodes = [node for node in root.iter() if self._local_name(node.tag) in _TASK_TAGS]
        activities: list[Activity] = []
        issues: list[ScheduleXmlIssue] = []
        existing_ids: set[str] = set()

        for order, task in enumerate(nodes):
            task_id = self._int(task, ("ID", "TaskID", "TaskId"))
            uid = self._int(task, ("UID", "Guid", "GUID"))
            is_summary = self._bool(task, ("Summary", "IsSummary"))
            name = self._clean(self._text(task, _NAME_TAGS))

            # Microsoft Project commonly includes the project summary as Task ID 0.
            if task_id == 0 and not name:
                continue

            plan_start = self._datetime(task, _START_TAGS)
            plan_finish = self._datetime(task, _FINISH_TAGS)

            if not is_summary:
                label = name or self._activity_id(task) or f"item {order + 1}"
                if not name:
                    issues.append(ScheduleXmlIssue(order, label, "Activity Name", "Activity Name is missing."))
                if plan_start is None:
                    issues.append(ScheduleXmlIssue(order, label, "Plan Start", "Plan Start is missing or invalid."))
                if plan_finish is None:
                    issues.append(ScheduleXmlIssue(order, label, "Plan Finish", "Plan Finish is missing or invalid."))
                if plan_start is not None and plan_finish is not None and plan_finish < plan_start:
                    issues.append(
                        ScheduleXmlIssue(
                            order,
                            label,
                            "Schedule Window",
                            "Plan Finish is earlier than Plan Start.",
                        )
                    )

            source_activity_id = self._activity_id(task)
            activity_id = ""
            if not is_summary:
                activity_id = ensure_activity_id(
                    source_activity_id,
                    task_id=task_id,
                    uid=uid,
                    source_order=order,
                    existing_ids=existing_ids,
                )
                existing_ids.add(activity_id)

            activities.append(
                Activity(
                    source_order=order,
                    task_id=task_id,
                    uid=uid,
                    activity_id=activity_id,
                    name=name,
                    wbs=self._text(task, _WBS_TAGS),
                    outline_level=self._int(task, ("OutlineLevel", "Level")) or (0 if is_summary else 1),
                    is_summary=is_summary,
                    plan_start=plan_start,
                    plan_finish=plan_finish,
                    actual_start=self._datetime(task, ("ActualStart",)),
                    actual_finish=self._datetime(task, ("ActualFinish",)),
                    percent_complete=self._float(task, ("PercentComplete",)),
                    physical_percent_complete=self._float(task, ("PhysicalPercentComplete",)),
                    total_slack_minutes=self._float(task, ("TotalSlack",)),
                    amount=self._float(task, ("Cost", "Amount")) if not is_summary else None,
                )
            )

        leaf_count = sum(not row.is_summary for row in activities)
        if leaf_count == 0:
            issues.append(ScheduleXmlIssue(0, project_name, "Activities", "No activities were found in the XML file."))
        if issues:
            raise ScheduleXmlValidationError(issues)
        return project_name, activities

    @staticmethod
    def _local_name(tag: str) -> str:
        return tag.rsplit("}", 1)[-1] if "}" in tag else tag.split(":", 1)[-1]

    def _direct_text(self, parent: ET.Element, tags: tuple[str, ...]) -> str:
        wanted = set(tags)
        for child in list(parent):
            if self._local_name(child.tag) in wanted and child.text:
                return self._clean(child.text)
        return ""

    def _text(self, parent: ET.Element, tags: tuple[str, ...]) -> str:
        wanted = set(tags)
        for child in list(parent):
            if self._local_name(child.tag) in wanted and child.text:
                return self._clean(child.text)
        return ""

    def _int(self, parent: ET.Element, tags: tuple[str, ...]) -> int | None:
        value = self._text(parent, tags)
        try:
            return int(float(value)) if value else None
        except ValueError:
            return None

    def _float(self, parent: ET.Element, tags: tuple[str, ...]) -> float | None:
        value = self._text(parent, tags)
        try:
            return float(value) if value else None
        except ValueError:
            return None

    def _bool(self, parent: ET.Element, tags: tuple[str, ...]) -> bool:
        value = self._text(parent, tags).strip().lower()
        return value in {"1", "true", "yes", "y"}

    def _datetime(self, parent: ET.Element, tags: tuple[str, ...]) -> datetime | None:
        value = self._text(parent, tags)
        if not value:
            return None
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            return None

    @staticmethod
    def _clean(value: str) -> str:
        previous = value
        for _ in range(3):
            current = html.unescape(previous)
            if current == previous:
                break
            previous = current
        return previous.strip()

    def _activity_id(self, task: ET.Element) -> str:
        value = self._text(task, _ACTIVITY_ID_TAGS)
        if value:
            return self._clean(value)

        for attribute in list(task):
            if self._local_name(attribute.tag) != "ExtendedAttribute":
                continue
            field_id = self._text(attribute, ("FieldID",))
            attribute_value = self._clean(self._text(attribute, ("Value",)))
            if field_id in P6_ACTIVITY_ID_FIELD_IDS and attribute_value:
                return attribute_value

        for attribute in list(task):
            if self._local_name(attribute.tag) != "ExtendedAttribute":
                continue
            label = f"{self._text(attribute, ('Alias',))} {self._text(attribute, ('FieldName',))}".lower()
            attribute_value = self._clean(self._text(attribute, ("Value",)))
            if attribute_value and (
                "activity id" in label or "activityid" in label or "activity code" in label
            ):
                return attribute_value
        return ""
