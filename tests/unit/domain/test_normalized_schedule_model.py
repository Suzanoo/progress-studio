from __future__ import annotations

from dataclasses import FrozenInstanceError
from datetime import datetime

import pytest

from progress_studio.domain import (
    NormalizedActivity,
    NormalizedProject,
    NormalizedSchedule,
    NormalizedWbs,
)


def test_n2_model_represents_same_canonical_wbs_for_msp_and_p6() -> None:
    project = NormalizedProject(project_id="007", project_name="NKC1R2")
    pile_wall = NormalizedWbs(
        source_order=3,
        wbs_code="2.1.1",
        wbs_name="Pile Wall",
        parent_wbs_code="2.1",
        outline_level=3,
    )
    activity = NormalizedActivity(
        source_order=4,
        activity_id="A1290",
        activity_name="Mobilization เครื่องจักรงานเสาเข็ม",
        wbs_code="2.1.1",
        outline_level=4,
        plan_start=datetime(2026, 10, 1, 8, 0),
        plan_finish=datetime(2026, 10, 7, 17, 0),
    )

    schedule = NormalizedSchedule(
        project=project,
        wbs=(pile_wall,),
        activities=(activity,),
    )

    assert schedule.project.project_id == "007"
    assert schedule.wbs[0].wbs_code == "2.1.1"
    assert schedule.activities[0].activity_id == "A1290"
    assert schedule.activities[0].wbs_code == "2.1.1"


def test_n2_amount_is_not_part_of_normalized_activity_contract() -> None:
    fields = NormalizedActivity.__dataclass_fields__
    assert "amount" not in fields
    assert "fixed_cost" not in fields


def test_n2_normalized_contract_is_immutable() -> None:
    project = NormalizedProject(project_id="007", project_name="NKC1R2")

    with pytest.raises(FrozenInstanceError):
        project.project_name = "changed"  # type: ignore[misc]
