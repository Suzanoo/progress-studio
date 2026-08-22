from __future__ import annotations

from pathlib import Path

from tests._paths import FIXTURES_ROOT

from progress_studio.infrastructure.schedule_xml import MspXmlAdapter, ScheduleXmlReader


FIXTURE = FIXTURES_ROOT / "xml" / "msp_n3.xml"


def _parent(code: str) -> str | None:
    code = (code or "").strip()
    if "." not in code:
        return None
    parent = code.rsplit(".", 1)[0].strip()
    return parent or None


def test_n4_normalized_msp_projection_matches_proven_reader_row_for_row() -> None:
    """Hard regression gate: N-3 must not reinterpret proven MSP reader data."""
    project_name, legacy_rows = ScheduleXmlReader().read(FIXTURE)
    normalized = MspXmlAdapter().normalize(FIXTURE)

    legacy_wbs = [row for row in legacy_rows if row.is_summary and row.wbs]
    legacy_activities = [row for row in legacy_rows if not row.is_summary]

    assert normalized.project.project_name == project_name
    assert len(normalized.wbs) == len(legacy_wbs)
    assert len(normalized.activities) == len(legacy_activities)

    for source, target in zip(legacy_wbs, normalized.wbs, strict=True):
        assert target.source_order == source.source_order
        assert target.wbs_code == source.wbs
        assert target.wbs_name == source.name
        assert target.parent_wbs_code == _parent(source.wbs)
        assert target.outline_level == source.outline_level

    for source, target in zip(legacy_activities, normalized.activities, strict=True):
        assert target.source_order == source.source_order
        assert target.activity_id == source.activity_id
        assert target.activity_name == source.name
        assert target.wbs_code == source.wbs
        assert target.outline_level == source.outline_level
        assert target.plan_start == source.plan_start
        assert target.plan_finish == source.plan_finish
        assert target.actual_start == source.actual_start
        assert target.actual_finish == source.actual_finish
        assert target.percent_complete == source.percent_complete
        assert target.physical_percent_complete == source.physical_percent_complete


def test_n4_project_schedule_window_matches_legacy_leaf_activity_window() -> None:
    _, legacy_rows = ScheduleXmlReader().read(FIXTURE)
    normalized = MspXmlAdapter().normalize(FIXTURE)
    leaves = [row for row in legacy_rows if not row.is_summary]

    starts = [row.plan_start for row in leaves if row.plan_start is not None]
    finishes = [row.plan_finish for row in leaves if row.plan_finish is not None]

    assert normalized.project.plan_start == min(starts)
    assert normalized.project.plan_finish == max(finishes)


def test_n4_msp_identity_and_hierarchy_contract_is_locked() -> None:
    normalized = MspXmlAdapter().normalize(FIXTURE)
    by_id = {row.activity_id: row for row in normalized.activities}
    wbs_by_code = {row.wbs_code: row for row in normalized.wbs}

    # P6 Activity ID exported through MSP Text1/ExtendedAttribute must survive intact.
    assert by_id["A1290"].wbs_code == "2.1.1"
    assert by_id["A1310"].wbs_code == "2.1.1.1"

    # WBS hierarchy used by workbook styling/outline must remain resolvable.
    assert wbs_by_code["2.1"].parent_wbs_code == "2"
    assert wbs_by_code["2.1.1"].parent_wbs_code == "2.1"
    assert wbs_by_code["2.1.1.1"].parent_wbs_code == "2.1.1"


def test_n4_normalized_schedule_does_not_import_source_cost() -> None:
    """Amount remains a downstream fake/mapping concern, not normalization."""
    normalized = MspXmlAdapter().normalize(FIXTURE)

    assert normalized.activities
    assert all(not hasattr(row, "amount") for row in normalized.activities)
    assert all(not hasattr(row, "fixed_cost") for row in normalized.activities)
