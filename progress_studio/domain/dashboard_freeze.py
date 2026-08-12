"""Stable Dashboard behavior contract.

LW-13 freezes the user-visible Dashboard/overlay behavior that was accepted in
LW-12.4.2.  This module is intentionally tiny: it gives future rebuild work a
single versioned contract to target without coupling the engine to worksheet
cell addresses beyond the established UI controls.

Changing any value in :data:`DASHBOARD_V1` is a behavior change and therefore
requires an explicit Dashboard contract version bump plus regression-test
updates.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class DashboardFreezeContract:
    version: str
    dashboard_view_cell: str
    dashboard_cutoff_cell: str
    traditional_cutoff_column: str
    weekly_cutoff_format: str
    monthly_cutoff_format: str
    plan_uses_full_timeline: bool
    actual_is_cutoff_masked: bool
    traditional_cutoffs_are_independent: bool
    traditional_overlay_has_legend: bool
    cutoff_line_has_label: bool
    cutoff_label_font_points: int


DASHBOARD_V1 = DashboardFreezeContract(
    version="dashboard-v1-stable",
    dashboard_view_cell="G5",
    dashboard_cutoff_cell="K5",
    traditional_cutoff_column="M",
    weekly_cutoff_format="dd/mm/yyyy",
    monthly_cutoff_format="mmm yyyy",
    plan_uses_full_timeline=True,
    actual_is_cutoff_masked=True,
    traditional_cutoffs_are_independent=True,
    traditional_overlay_has_legend=False,
    cutoff_line_has_label=True,
    cutoff_label_font_points=10,
)
