from __future__ import annotations

from pathlib import Path

from progress_studio.infrastructure.schedule_xml import ScheduleXmlFormat, ScheduleXmlFormatDetector


MSP_XML = """<?xml version='1.0' encoding='utf-8'?>
<Project xmlns='http://schemas.microsoft.com/project'>
  <Name>Sample MSP</Name>
</Project>
"""

P6_XML = """<?xml version='1.0' encoding='utf-8'?>
<APIBusinessObjects xmlns='http://xmlns.oracle.com/Primavera/P6Professional/V24.12/API/BusinessObjects'>
  <Project><Id>007</Id></Project>
</APIBusinessObjects>
"""

UNKNOWN_XML = """<?xml version='1.0' encoding='utf-8'?>
<Schedule><ProjectName>Generic</ProjectName></Schedule>
"""


def _write(tmp_path: Path, name: str, payload: str) -> Path:
    path = tmp_path / name
    path.write_text(payload, encoding="utf-8")
    return path


def test_detects_microsoft_project_xml(tmp_path: Path) -> None:
    detector = ScheduleXmlFormatDetector()
    assert detector.detect(_write(tmp_path, "msp.xml", MSP_XML)) is ScheduleXmlFormat.MSP_XML


def test_detects_primavera_p6_api_xml(tmp_path: Path) -> None:
    detector = ScheduleXmlFormatDetector()
    assert detector.detect(_write(tmp_path, "p6.xml", P6_XML)) is ScheduleXmlFormat.P6_XML


def test_unknown_xml_is_not_misclassified(tmp_path: Path) -> None:
    detector = ScheduleXmlFormatDetector()
    assert detector.detect(_write(tmp_path, "generic.xml", UNKNOWN_XML)) is ScheduleXmlFormat.UNKNOWN


def test_malformed_or_missing_xml_is_unknown(tmp_path: Path) -> None:
    detector = ScheduleXmlFormatDetector()
    assert detector.detect(_write(tmp_path, "broken.xml", "<Project>")) is ScheduleXmlFormat.UNKNOWN
    assert detector.detect(tmp_path / "missing.xml") is ScheduleXmlFormat.UNKNOWN
