from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from progress_studio.domain import (
    NormalizedActivity,
    NormalizedProject,
    NormalizedSchedule,
    NormalizedWbs,
)

from .format_detector import ScheduleXmlFormat, ScheduleXmlFormatDetector


@dataclass(frozen=True)
class _P6WbsSource:
    xml_order: int
    object_id: str
    code: str
    name: str
    parent_object_id: str | None
    sequence_number: int | None


class P6XmlAdapter:
    """Normalize Primavera P6 API XML into the source-neutral schedule model.

    P6 stores WBS nodes independently from activities.  Activities reference a
    WBS by ``WBSObjectId``; each WBS node stores only its *local* ``Code`` plus
    ``ParentObjectId``.  This adapter resolves that object graph and reconstructs
    the canonical WBS path expected by Progress Studio (for example ``2.1.1``),
    deliberately excluding the P6 project prefix (for example ``007``).

    Amount/cost fields are intentionally ignored.  Fake Amount remains a later
    Create Progress concern, matching the existing MSP workflow.
    """

    def __init__(self, detector: ScheduleXmlFormatDetector | None = None) -> None:
        self._detector = detector or ScheduleXmlFormatDetector()

    def normalize(self, xml_file: Path) -> NormalizedSchedule:
        detected = self._detector.detect(xml_file)
        if detected is not ScheduleXmlFormat.P6_XML:
            raise ValueError(f"P6XmlAdapter requires Primavera P6 XML; detected {detected.value}.")

        try:
            root = ET.parse(xml_file).getroot()
        except (ET.ParseError, OSError) as exc:  # detector normally catches this first
            raise ValueError(f"Invalid Primavera P6 XML: {exc}") from exc

        namespace = self._namespace(root.tag)
        project = self._first_direct_child(root, "Project", namespace)
        if project is None:
            raise ValueError("Primavera P6 XML does not contain a Project element.")

        separator = self._text(project, "WBSCodeSeparator", namespace) or "."
        project_id = self._text(project, "Id", namespace) or None
        project_name = self._text(project, "Name", namespace) or xml_file.stem

        source_wbs = self._read_wbs(project, namespace)
        by_object_id = {node.object_id: node for node in source_wbs}
        code_cache: dict[str, str] = {}
        level_cache: dict[str, int] = {}

        def resolve_code(object_id: str, trail: tuple[str, ...] = ()) -> str:
            if object_id in code_cache:
                return code_cache[object_id]
            if object_id in trail:
                raise ValueError(f"P6 WBS hierarchy contains a cycle at ObjectId {object_id}.")
            node = by_object_id.get(object_id)
            if node is None:
                raise ValueError(f"P6 WBS ObjectId {object_id} could not be resolved.")
            if node.parent_object_id:
                parent_code = resolve_code(node.parent_object_id, trail + (object_id,))
                code = f"{parent_code}{separator}{node.code}" if parent_code else node.code
            else:
                code = node.code
            code_cache[object_id] = code
            return code

        def resolve_level(object_id: str, trail: tuple[str, ...] = ()) -> int:
            if object_id in level_cache:
                return level_cache[object_id]
            if object_id in trail:
                raise ValueError(f"P6 WBS hierarchy contains a cycle at ObjectId {object_id}.")
            node = by_object_id.get(object_id)
            if node is None:
                raise ValueError(f"P6 WBS ObjectId {object_id} could not be resolved.")
            level = (
                resolve_level(node.parent_object_id, trail + (object_id,)) + 1
                if node.parent_object_id
                else 1
            )
            level_cache[object_id] = level
            return level

        ordered_wbs = sorted(
            source_wbs,
            key=lambda row: (
                row.sequence_number is None,
                row.sequence_number if row.sequence_number is not None else row.xml_order,
                row.xml_order,
            ),
        )
        normalized_wbs: list[NormalizedWbs] = []
        for source_order, node in enumerate(ordered_wbs):
            parent_code = resolve_code(node.parent_object_id) if node.parent_object_id else None
            normalized_wbs.append(
                NormalizedWbs(
                    source_order=source_order,
                    wbs_code=resolve_code(node.object_id),
                    wbs_name=node.name,
                    parent_wbs_code=parent_code,
                    outline_level=resolve_level(node.object_id),
                )
            )

        normalized_activities: list[NormalizedActivity] = []
        for source_order, activity in enumerate(self._direct_children(project, "Activity", namespace)):
            activity_id = self._text(activity, "Id", namespace)
            wbs_object_id = self._text(activity, "WBSObjectId", namespace)
            if not activity_id:
                raise ValueError(f"P6 activity at source order {source_order} has no Activity Id.")
            if not wbs_object_id:
                raise ValueError(f"P6 activity {activity_id} has no WBSObjectId.")
            if wbs_object_id not in by_object_id:
                raise ValueError(
                    f"P6 activity {activity_id} references unknown WBSObjectId {wbs_object_id}."
                )

            normalized_activities.append(
                NormalizedActivity(
                    source_order=source_order,
                    activity_id=activity_id,
                    activity_name=self._text(activity, "Name", namespace),
                    wbs_code=resolve_code(wbs_object_id),
                    outline_level=resolve_level(wbs_object_id) + 1,
                    plan_start=self._datetime(activity, "PlannedStartDate", namespace),
                    plan_finish=self._datetime(activity, "PlannedFinishDate", namespace),
                    actual_start=self._datetime(activity, "ActualStartDate", namespace),
                    actual_finish=self._datetime(activity, "ActualFinishDate", namespace),
                    percent_complete=self._float(activity, "PercentComplete", namespace),
                    physical_percent_complete=self._float(
                        activity, "PhysicalPercentComplete", namespace
                    ),
                )
            )

        plan_starts = [row.plan_start for row in normalized_activities if row.plan_start is not None]
        plan_finishes = [row.plan_finish for row in normalized_activities if row.plan_finish is not None]
        normalized_project = NormalizedProject(
            project_id=project_id,
            project_name=project_name,
            plan_start=min(plan_starts) if plan_starts else None,
            plan_finish=max(plan_finishes) if plan_finishes else None,
        )

        return NormalizedSchedule(
            project=normalized_project,
            wbs=tuple(normalized_wbs),
            activities=tuple(normalized_activities),
        )

    def _read_wbs(self, project: ET.Element, namespace: str) -> list[_P6WbsSource]:
        rows: list[_P6WbsSource] = []
        for xml_order, element in enumerate(self._direct_children(project, "WBS", namespace)):
            object_id = self._text(element, "ObjectId", namespace)
            code = self._text(element, "Code", namespace)
            if not object_id or not code:
                raise ValueError(f"P6 WBS at source order {xml_order} is missing ObjectId or Code.")
            rows.append(
                _P6WbsSource(
                    xml_order=xml_order,
                    object_id=object_id,
                    code=code,
                    name=self._text(element, "Name", namespace),
                    parent_object_id=self._text(element, "ParentObjectId", namespace) or None,
                    sequence_number=self._int(element, "SequenceNumber", namespace),
                )
            )
        return rows

    @classmethod
    def _first_direct_child(
        cls, parent: ET.Element, local_name: str, namespace: str
    ) -> ET.Element | None:
        for child in list(parent):
            if cls._local_name(child.tag) == local_name and cls._namespace(child.tag) == namespace:
                return child
        return None

    @classmethod
    def _direct_children(
        cls, parent: ET.Element, local_name: str, namespace: str
    ) -> list[ET.Element]:
        return [
            child
            for child in list(parent)
            if cls._local_name(child.tag) == local_name and cls._namespace(child.tag) == namespace
        ]

    @classmethod
    def _text(cls, parent: ET.Element, local_name: str, namespace: str) -> str:
        child = cls._first_direct_child(parent, local_name, namespace)
        if child is None or not child.text:
            return ""
        return cls._clean(child.text)

    @classmethod
    def _datetime(cls, parent: ET.Element, local_name: str, namespace: str) -> datetime | None:
        value = cls._text(parent, local_name, namespace)
        if not value:
            return None
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            return None

    @classmethod
    def _float(cls, parent: ET.Element, local_name: str, namespace: str) -> float | None:
        value = cls._text(parent, local_name, namespace)
        try:
            return float(value) if value else None
        except ValueError:
            return None

    @classmethod
    def _int(cls, parent: ET.Element, local_name: str, namespace: str) -> int | None:
        value = cls._text(parent, local_name, namespace)
        try:
            return int(float(value)) if value else None
        except ValueError:
            return None

    @staticmethod
    def _namespace(tag: str) -> str:
        return tag[1:].split("}", 1)[0] if tag.startswith("{") and "}" in tag else ""

    @staticmethod
    def _local_name(tag: str) -> str:
        return tag.rsplit("}", 1)[-1] if "}" in tag else tag.split(":", 1)[-1]

    @staticmethod
    def _clean(value: str) -> str:
        previous = value
        for _ in range(3):
            current = html.unescape(previous)
            if current == previous:
                break
            previous = current
        return previous.strip()
