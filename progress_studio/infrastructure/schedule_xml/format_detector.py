from __future__ import annotations

import xml.etree.ElementTree as ET
from enum import Enum
from pathlib import Path


MS_PROJECT_NAMESPACE = "http://schemas.microsoft.com/project"
P6_NAMESPACE_MARKER = "xmlns.oracle.com/Primavera/P6Professional"


class ScheduleXmlFormat(str, Enum):
    """Schedule XML formats understood by the import boundary.

    Detection is deliberately separate from parsing.  N-1 only establishes a
    stable source classification contract; adapters are introduced in later
    milestones.
    """

    MSP_XML = "msp_xml"
    P6_XML = "p6_xml"
    UNKNOWN = "unknown"


class ScheduleXmlFormatDetector:
    """Classify a schedule XML document without interpreting schedule data."""

    def detect(self, xml_file: Path) -> ScheduleXmlFormat:
        try:
            root = ET.parse(xml_file).getroot()
        except (ET.ParseError, OSError):
            return ScheduleXmlFormat.UNKNOWN
        return self.detect_root(root)

    @classmethod
    def detect_root(cls, root: ET.Element) -> ScheduleXmlFormat:
        namespace, local_name = cls._split_tag(root.tag)

        if local_name == "Project" and namespace == MS_PROJECT_NAMESPACE:
            return ScheduleXmlFormat.MSP_XML

        if local_name == "APIBusinessObjects" and P6_NAMESPACE_MARKER in namespace:
            return ScheduleXmlFormat.P6_XML

        return ScheduleXmlFormat.UNKNOWN

    @staticmethod
    def _split_tag(tag: str) -> tuple[str, str]:
        if tag.startswith("{") and "}" in tag:
            namespace, local_name = tag[1:].split("}", 1)
            return namespace, local_name
        if ":" in tag:
            return "", tag.split(":", 1)[-1]
        return "", tag
