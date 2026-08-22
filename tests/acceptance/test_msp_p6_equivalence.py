from __future__ import annotations

from pathlib import Path

from tests._paths import FIXTURES_ROOT

from progress_studio.infrastructure.schedule_xml import (
    MspXmlAdapter,
    NormalizedScheduleXmlReader,
    P6XmlAdapter,
)

FIXTURES = FIXTURES_ROOT / "xml"
MSP = FIXTURES / "msp_n8.xml"
P6 = FIXTURES / "p6_n8.xml"


def _wbs_projection(schedule):
    return {
        row.wbs_code: (row.wbs_name, row.parent_wbs_code, row.outline_level)
        for row in schedule.wbs
    }


def _activity_projection(schedule):
    return {
        row.activity_id: (
            row.activity_name,
            row.wbs_code,
            row.outline_level,
            row.plan_start,
            row.plan_finish,
        )
        for row in schedule.activities
    }


def _legacy_projection(rows):
    wbs = {
        row.wbs: (row.name, row.outline_level)
        for row in rows
        if row.is_summary and row.wbs
    }
    activities = {
        row.activity_id: (
            row.name,
            row.wbs,
            row.outline_level,
            row.plan_start,
            row.plan_finish,
            row.amount,
        )
        for row in rows
        if not row.is_summary
    }
    return wbs, activities


def test_n8_adapters_produce_equivalent_canonical_schedule_structure() -> None:
    """Freeze the shared schedule contract across MSP and P6 source formats.

    The fixtures intentionally model the same WBS/activity baseline while
    retaining source-specific metadata differences (for example project name
    and some Actual fields).  N-8 compares only the canonical fields consumed
    by Create Progress.
    """
    msp = MspXmlAdapter().normalize(MSP)
    p6 = P6XmlAdapter().normalize(P6)

    assert _wbs_projection(msp) == _wbs_projection(p6)
    assert _activity_projection(msp) == _activity_projection(p6)

    assert msp.project.plan_start == p6.project.plan_start
    assert msp.project.plan_finish == p6.project.plan_finish


def test_n8_activity_identity_and_wbs_assignment_are_source_independent() -> None:
    msp = MspXmlAdapter().normalize(MSP)
    p6 = P6XmlAdapter().normalize(P6)

    msp_by_id = {row.activity_id: row for row in msp.activities}
    p6_by_id = {row.activity_id: row for row in p6.activities}

    assert set(msp_by_id) == set(p6_by_id) == {"A1290", "A1310"}
    for activity_id in sorted(msp_by_id):
        assert msp_by_id[activity_id].wbs_code == p6_by_id[activity_id].wbs_code
        assert not activity_id.startswith("ACT-")


def test_n8_production_reader_emits_equivalent_create_progress_rows() -> None:
    """Freeze the production bridge, not only the adapter implementation."""
    reader = NormalizedScheduleXmlReader()
    _, msp_rows = reader.read(MSP)
    _, p6_rows = reader.read(P6)

    msp_wbs, msp_activities = _legacy_projection(msp_rows)
    p6_wbs, p6_activities = _legacy_projection(p6_rows)

    assert msp_wbs == p6_wbs
    assert msp_activities == p6_activities

    # Normalization owns schedule structure only.  Fake Amount is still a
    # downstream Create Progress concern for both source formats.
    assert all(values[-1] is None for values in msp_activities.values())
    assert all(values[-1] is None for values in p6_activities.values())


def test_n8_project_prefix_is_metadata_not_canonical_wbs() -> None:
    msp = MspXmlAdapter().normalize(MSP)
    p6 = P6XmlAdapter().normalize(P6)

    assert p6.project.project_id == "007"
    assert all(not row.wbs_code.startswith("007.") for row in p6.wbs)
    assert {row.wbs_code for row in msp.wbs} == {row.wbs_code for row in p6.wbs}
