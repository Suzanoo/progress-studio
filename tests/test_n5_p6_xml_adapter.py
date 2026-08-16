from __future__ import annotations

from pathlib import Path

import pytest

from progress_studio.infrastructure.schedule_xml import P6XmlAdapter

FIXTURES = Path(__file__).parent / "fixtures" / "xml"


def test_p6_adapter_reconstructs_wbs_object_graph_into_canonical_codes() -> None:
    schedule = P6XmlAdapter().normalize(FIXTURES / "p6_n5.xml")

    assert schedule.project.project_id == "007"
    assert schedule.project.project_name == "N5 P6 Fixture"
    assert schedule.project.plan_start.isoformat() == "2026-10-01T08:00:00"
    assert schedule.project.plan_finish.isoformat() == "2026-11-16T17:00:00"

    assert [(row.wbs_code, row.parent_wbs_code, row.outline_level) for row in schedule.wbs] == [
        ("2", None, 1),
        ("2.1", "2", 2),
        ("2.1.1", "2.1", 3),
        ("2.1.1.1", "2.1.1", 4),
    ]


def test_p6_adapter_preserves_activity_id_and_resolves_activity_wbs() -> None:
    schedule = P6XmlAdapter().normalize(FIXTURES / "p6_n5.xml")

    assert [row.activity_id for row in schedule.activities] == ["A1290", "A1310"]
    assert [row.wbs_code for row in schedule.activities] == ["2.1.1", "2.1.1.1"]
    assert [row.outline_level for row in schedule.activities] == [4, 5]
    assert schedule.activities[0].percent_complete == 25.0
    assert schedule.activities[0].physical_percent_complete == 20.0
    assert schedule.activities[1].actual_start.isoformat() == "2026-10-08T08:00:00"


def test_p6_adapter_does_not_import_schedule_cost_as_progress_amount() -> None:
    schedule = P6XmlAdapter().normalize(FIXTURES / "p6_n5.xml")
    activity = schedule.activities[0]

    assert not hasattr(activity, "amount")
    assert not hasattr(activity, "fixed_cost")
    assert not hasattr(activity, "total_cost")


def test_p6_adapter_rejects_msp_xml_before_interpretation() -> None:
    with pytest.raises(ValueError, match="requires Primavera P6 XML"):
        P6XmlAdapter().normalize(FIXTURES / "msp_n3.xml")


def test_p6_adapter_never_fabricates_wbs_when_reference_is_unknown() -> None:
    with pytest.raises(ValueError, match="unknown WBSObjectId 999999"):
        P6XmlAdapter().normalize(FIXTURES / "p6_n5_unknown_wbs.xml")
