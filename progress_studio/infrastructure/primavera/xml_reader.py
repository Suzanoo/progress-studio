from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

from progress_studio.domain.activity_id import ensure_activity_id
from progress_studio.domain import Activity

P6_ACTIVITY_ID_FIELD_IDS = {"188743731"}


class PrimaveraXmlReader:
    def read(self, xml_file: Path) -> tuple[str, list[Activity]]:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        namespace = self._namespace(root)
        project_name = self._text(root, "Name", namespace, xml_file.stem)
        activities: list[Activity] = []

        for order, task in enumerate(root.findall("msp:Tasks/msp:Task", namespace)):
            task_id = self._int(task, "ID", namespace)
            is_summary = self._int(task, "Summary", namespace) == 1
            name = self._clean(self._text(task, "Name", namespace))
            if task_id == 0 and not name:
                continue
            uid = self._int(task, "UID", namespace)
            activities.append(
                Activity(
                    source_order=order,
                    task_id=task_id,
                    uid=uid,
                    activity_id=(
                        ensure_activity_id(
                            self._activity_id(task, namespace),
                            task_id=task_id,
                            uid=uid,
                            source_order=order,
                        )
                        if not is_summary
                        else ""
                    ),
                    name=name,
                    wbs=self._text(task, "WBS", namespace),
                    outline_level=self._int(task, "OutlineLevel", namespace) or 0,
                    is_summary=is_summary,
                    plan_start=self._datetime(task, "Start", namespace),
                    plan_finish=self._datetime(task, "Finish", namespace),
                    actual_start=self._datetime(task, "ActualStart", namespace),
                    actual_finish=self._datetime(task, "ActualFinish", namespace),
                    percent_complete=self._float(task, "PercentComplete", namespace),
                    physical_percent_complete=self._float(task, "PhysicalPercentComplete", namespace),
                    total_slack_minutes=self._float(task, "TotalSlack", namespace),
                    amount=self._float(task, "Cost", namespace) if not is_summary else None,
                )
            )
        return project_name, activities

    @staticmethod
    def _namespace(root: ET.Element) -> dict[str, str]:
        uri = "http://schemas.microsoft.com/project/2007"
        if root.tag.startswith("{") and "}" in root.tag:
            uri = root.tag[1:].split("}", 1)[0]
        return {"msp": uri}

    @staticmethod
    def _text(parent: ET.Element, tag: str, ns: dict[str, str], default: str = "") -> str:
        node = parent.find(f"msp:{tag}", ns)
        return html.unescape(node.text).strip() if node is not None and node.text else default

    def _int(self, parent: ET.Element, tag: str, ns: dict[str, str]) -> int | None:
        value = self._text(parent, tag, ns)
        try:
            return int(float(value)) if value else None
        except ValueError:
            return None

    def _float(self, parent: ET.Element, tag: str, ns: dict[str, str]) -> float | None:
        value = self._text(parent, tag, ns)
        try:
            return float(value) if value else None
        except ValueError:
            return None

    def _datetime(self, parent: ET.Element, tag: str, ns: dict[str, str]) -> datetime | None:
        value = self._text(parent, tag, ns)
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None) if value else None
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

    def _activity_id(self, task: ET.Element, ns: dict[str, str]) -> str:
        for tag in ("ActivityID", "ActivityId", "ActivityCode", "ActivityCodeValue"):
            value = self._text(task, tag, ns)
            if value:
                return self._clean(value)
        for attribute in task.findall("msp:ExtendedAttribute", ns):
            field_id = self._text(attribute, "FieldID", ns)
            value = self._clean(self._text(attribute, "Value", ns))
            if field_id in P6_ACTIVITY_ID_FIELD_IDS and value:
                return value
        for attribute in task.findall("msp:ExtendedAttribute", ns):
            label = f"{self._text(attribute, 'Alias', ns)} {self._text(attribute, 'FieldName', ns)}".lower()
            value = self._clean(self._text(attribute, "Value", ns))
            if value and ("activity id" in label or "activityid" in label or "activity code" in label):
                return value
        return ""
