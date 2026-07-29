"""Backward-compatible import for the former Primavera-specific reader."""

from progress_studio.infrastructure.schedule_xml import ScheduleXmlReader


class PrimaveraXmlReader(ScheduleXmlReader):
    pass
