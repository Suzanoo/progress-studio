from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TimescalePalette:
    """Colors used by conditional formatting in the main-sheet timescale."""

    activity_plan_fill: str = "DDEBF7"
    activity_actual_fill: str = "E2F0D9"
    project_plan_fill: str = "1F4E78"
    project_actual_fill: str = "375623"
    wbs_level_1_plan_fill: str = "2F5597"
    wbs_level_1_actual_fill: str = "548235"
    wbs_level_2_plan_fill: str = "5B9BD5"
    wbs_level_2_actual_fill: str = "70AD47"
    scurve_plan_fill: str = "9DC3E6"
    scurve_actual_fill: str = "A9D18E"
    scurve_acc_fill: str = "D9EAD3"


@dataclass(frozen=True)
class ActivityDataPalette:
    """Colors used only by the Activity Data section of the main sheet."""

    wbs_level_1_fill: str = "F4B183"
    wbs_level_2_fill: str = "F8CBAD"
    wbs_level_3_fill: str = "FADBC8"
    wbs_level_4_fill: str = "FCE8DE"
    font_color: str = "000000"


DEFAULT_TIMESCALE_PALETTE = TimescalePalette()
DEFAULT_ACTIVITY_DATA_PALETTE = ActivityDataPalette()
