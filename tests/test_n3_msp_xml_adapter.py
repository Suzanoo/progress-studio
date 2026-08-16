from __future__ import annotations

from pathlib import Path

import pytest

from progress_studio.infrastructure.schedule_xml import MspXmlAdapter

FIXTURES = Path(__file__).parent / "fixtures" / "xml"


def test_msp_adapter_normalizes_proven_reader_output() -> None:
    schedule = MspXmlAdapter().normalize(FIXTURES / "msp_n3.xml")

    assert schedule.project.project_id is None
    assert schedule.project.project_name == "N3 MSP Fixture"
    assert schedule.project.plan_start.isoformat() == "2026-10-01T08:00:00"
    assert schedule.project.plan_finish.isoformat() == "2026-11-16T17:00:00"

    assert [(row.wbs_code, row.parent_wbs_code, row.outline_level) for row in schedule.wbs] == [
        ("2", None, 1),
        ("2.1", "2", 2),
        ("2.1.1", "2.1", 3),
        ("2.1.1.1", "2.1.1", 4),
    ]

    assert [row.activity_id for row in schedule.activities] == ["A1290", "A1310"]
    assert [row.wbs_code for row in schedule.activities] == ["2.1.1", "2.1.1.1"]
    assert [row.outline_level for row in schedule.activities] == [4, 5]
    assert schedule.activities[0].activity_name == "Mobilization"
    assert schedule.activities[0].percent_complete == 25.0
    assert schedule.activities[0].physical_percent_complete == 20.0


def test_msp_adapter_preserves_text1_activity_identity_not_task_id() -> None:
    schedule = MspXmlAdapter().normalize(FIXTURES / "msp_n3.xml")
    first = schedule.activities[0]
    assert first.activity_id == "A1290"
    assert first.activity_id != "4"


def test_msp_adapter_rejects_non_msp_xml_before_reader_interpretation() -> None:
    with pytest.raises(ValueError, match="requires Microsoft Project XML"):
        MspXmlAdapter().normalize(FIXTURES / "generic_flat.xml")
