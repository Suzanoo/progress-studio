from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path

from tests._paths import FIXTURES_ROOT

import pytest

from progress_studio.domain import (
    NormalizedActivity,
    NormalizedProject,
    NormalizedSchedule,
    NormalizedScheduleValidationError,
    NormalizedScheduleValidator,
    NormalizedWbs,
)
from progress_studio.infrastructure.schedule_xml import MspXmlAdapter, P6XmlAdapter

FIXTURES = FIXTURES_ROOT / "xml"


def _valid_schedule() -> NormalizedSchedule:
    return NormalizedSchedule(
        project=NormalizedProject(
            project_id="007",
            project_name="N6 Fixture",
            plan_start=datetime(2026, 10, 1, 8),
            plan_finish=datetime(2026, 10, 31, 17),
        ),
        wbs=(
            NormalizedWbs(0, "2", "Structure", None, 1),
            NormalizedWbs(1, "2.1", "Pile Work", "2", 2),
        ),
        activities=(
            NormalizedActivity(
                source_order=2,
                activity_id="A1000",
                activity_name="Mobilization",
                wbs_code="2.1",
                outline_level=3,
                plan_start=datetime(2026, 10, 1, 8),
                plan_finish=datetime(2026, 10, 7, 17),
            ),
        ),
    )


def _fields(exc: NormalizedScheduleValidationError) -> set[str]:
    return {issue.field for issue in exc.issues}


def test_n6_real_msp_and_p6_normalized_schedules_pass_same_validator() -> None:
    validator = NormalizedScheduleValidator()

    msp = MspXmlAdapter().normalize(FIXTURES / "msp_n3.xml")
    p6 = P6XmlAdapter().normalize(FIXTURES / "p6_n5.xml")

    assert validator.validate(msp) is msp
    assert validator.validate(p6) is p6


def test_n6_rejects_blank_and_duplicate_activity_ids_without_fabricating_ids() -> None:
    schedule = _valid_schedule()
    first = replace(schedule.activities[0], activity_id="")
    second = replace(schedule.activities[0], source_order=3, activity_id="A1000")
    third = replace(schedule.activities[0], source_order=4, activity_id="A1000")
    broken = replace(schedule, activities=(first, second, third))

    with pytest.raises(NormalizedScheduleValidationError) as caught:
        NormalizedScheduleValidator().validate(broken)

    assert "Activity ID" in _fields(caught.value)
    assert "fabricate" in str(caught.value)
    assert [row.activity_id for row in broken.activities] == ["", "A1000", "A1000"]


def test_n6_rejects_activity_wbs_that_does_not_resolve() -> None:
    schedule = _valid_schedule()
    broken = replace(
        schedule,
        activities=(replace(schedule.activities[0], wbs_code="9.9"),),
    )

    with pytest.raises(NormalizedScheduleValidationError) as caught:
        NormalizedScheduleValidator().validate(broken)

    assert "Activity WBS" in _fields(caught.value)
    assert "unknown WBS code '9.9'" in str(caught.value)


def test_n6_rejects_missing_wbs_parent_and_wbs_cycles() -> None:
    schedule = _valid_schedule()
    missing_parent = replace(
        schedule,
        wbs=(replace(schedule.wbs[0], parent_wbs_code="9"), schedule.wbs[1]),
    )
    with pytest.raises(NormalizedScheduleValidationError) as missing:
        NormalizedScheduleValidator().validate(missing_parent)
    assert "WBS Parent" in _fields(missing.value)

    cycle = replace(
        schedule,
        wbs=(
            replace(schedule.wbs[0], parent_wbs_code="2.1"),
            replace(schedule.wbs[1], parent_wbs_code="2"),
        ),
    )
    with pytest.raises(NormalizedScheduleValidationError) as cyclic:
        NormalizedScheduleValidator().validate(cycle)
    assert "cycle" in str(cyclic.value)


def test_n6_rejects_missing_or_reversed_plan_dates() -> None:
    schedule = _valid_schedule()
    missing_start = replace(schedule.activities[0], plan_start=None)
    reversed_dates = replace(
        schedule.activities[0],
        source_order=3,
        activity_id="A1010",
        plan_start=datetime(2026, 11, 2, 8),
        plan_finish=datetime(2026, 11, 1, 17),
    )
    broken = replace(schedule, activities=(missing_start, reversed_dates))

    with pytest.raises(NormalizedScheduleValidationError) as caught:
        NormalizedScheduleValidator().validate(broken)

    fields = _fields(caught.value)
    assert "Plan Start" in fields
    assert "Plan Dates" in fields


def test_n6_rejects_empty_activity_set() -> None:
    schedule = replace(_valid_schedule(), activities=())

    with pytest.raises(NormalizedScheduleValidationError) as caught:
        NormalizedScheduleValidator().validate(schedule)

    assert _fields(caught.value) == {"Activities"}


def test_n6_does_not_know_about_amount_or_timescale_display_margin() -> None:
    schedule = _valid_schedule()
    validated = NormalizedScheduleValidator().validate(schedule)

    assert not hasattr(validated, "display_start")
    assert not hasattr(validated, "display_finish")
    assert all(not hasattr(row, "amount") for row in validated.activities)
    assert all(not hasattr(row, "fixed_cost") for row in validated.activities)
