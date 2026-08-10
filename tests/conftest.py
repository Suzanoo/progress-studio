from __future__ import annotations

from pathlib import Path

import pytest


# MS-TEST1 ownership tiers.
#
# FROZEN:
# Stable product contracts. They still run at release/full-suite gates, but are
# intentionally excluded from routine milestone regression runs.
#
# ACTIVE:
# Workbook/dashboard/payment/rebuild areas still changing in current milestones.
#
# SMOKE:
# Small cross-section of high-value contracts intended for the fastest local gate.
FROZEN_FILES = {
    "test_boq_mapping_service.py",
    "test_desktop_phase2.py",
    "test_entrypoints.py",
    "test_generic_xml_import.py",
    "test_mapping_store.py",
    "test_ms10_mapped_to_readability.py",
    "test_ms111_unified_working_tree.py",
    "test_ms112_core_tree_editing.py",
    "test_ms113_recursive_working_tree.py",
    "test_ms114_main_rebuild_foundation.py",
    # Frozen before LW refactor: preserves pre-LW shell/export workspace contracts.
    "test_ms91_production_ui.py",
    "test_ms93_professional_desktop_ui.py",
    "test_ms_rb6_rebuild_workspace_ui.py",
    "test_ms3_mapping_engine.py",
    "test_ms4_share_allocation.py",
    "test_ms5_persistent_session.py",
    "test_ms63_scope.py",
    "test_ms6_release_architecture.py",
    "test_ms7_release.py",
    "test_ms7_workspace_ux.py",
    "test_ms8_architecture_cleanup.py",
    "test_ms92_focus_mapping.py",
    "test_ms94_mapping_productivity.py",
    "test_ms95_supplemental_structure.py",
    "test_ms961_unified_progress_tree.py",
    "test_v202_bugfixes.py",
}

ACTIVE_FILES = {
    # Live Workbook refactor ownership begins here.
    "test_lw0_rebuild_export_ux.py",
    "test_lw1_rebuild_reader_contract.py",
    "test_lw2_main_dataset_parser.py",
    "test_lw3_direct_activity_deriver.py",
    "test_lw4_tiny_progress_cache.py",
    "test_lw5_live_dashboard_contract.py",
    "test_lw6_monthly_engine.py",
    "test_ms_p120_activity_status_focus.py",
    "test_ms_rb721_dashboard_interactive_protection.py",
    "test_ms_rb72_lightweight_protection.py",
    "test_ms_rb711_hybrid_progress_contract.py",
    "test_ms_rb71_final_sheet_visibility.py",
    "test_ms_test1_test_suite_tiering.py",
    "test_ms115_workbook_generation_engine.py",
    "test_ms116_generation_progress_dialog.py",
    "test_ms6_workbook_export.py",
    "test_ms_p110_dashboard_embedded_icons.py",
    "test_ms_p111_dashboard_control_table.py",
    "test_ms_p113_filter_ui.py",
    "test_ms_p117_monthly_main_view.py",
    "test_ms_p119_snapshot_progress_table.py",
    "test_ms_p13_activity_data_theme.py",
    "test_ms_p15_actual_amount.py",
    "test_ms_p16_excel_dashboard.py",
    "test_ms_p17_dashboard_at_progress_stage.py",
    "test_ms_p19_dashboard_ui_behavior.py",
    "test_ms_pay1_payment_snapshot.py",
    "test_ms_pay2_payment_workflow.py",
    "test_ms_pay3_pay5_payment_position_engine.py",
    "test_ms_pay6_payment_line_renderer.py",
    "test_ms_pay7_embedded_payment_workflow.py",
    "test_ms_r2_rebuild_from_edited_workbook.py",
    "test_ms_rb1_rebuild_core_contract.py",
    "test_ms_rb2_progress_rebuild_engine.py",
    "test_ms_rb3_dashboard_progress_contract.py",
    "test_ms_rb3_snapshot_performance.py",
    "test_ms_rb4_payment_only_rebuild.py",
    "test_ms_rb5_payment_collision_lanes.py",
}

SMOKE_NODEIDS = {
    "test_lw0_rebuild_export_ux.py",
    "test_lw1_rebuild_reader_contract.py",
    "test_entrypoints.py",
    "test_ms115_workbook_generation_engine.py",
    "test_ms_p119_snapshot_progress_table.py",
    "test_ms_pay3_pay5_payment_position_engine.py",
    "test_ms_pay6_payment_line_renderer.py::test_ms_pay65_default_render_includes_every_populated_payment",
    "test_ms_rb1_rebuild_core_contract.py",
    "test_ms_rb2_progress_rebuild_engine.py::test_rb2_rebuild_progress_replaces_only_progress_owned_sheets",
    "test_ms_rb3_snapshot_performance.py",
    "test_ms_rb4_payment_only_rebuild.py::test_rb4_rebuild_payment_replaces_payment_only",
    "test_ms_rb5_payment_collision_lanes.py",
}


def _matches_smoke(item: pytest.Item) -> bool:
    short = item.nodeid.replace("\\", "/").split("tests/", 1)[-1]
    filename = Path(str(item.fspath)).name
    return filename in SMOKE_NODEIDS or short in SMOKE_NODEIDS


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    unknown: set[str] = set()

    for item in items:
        filename = Path(str(item.fspath)).name

        # Every test is part of the release gate.
        item.add_marker(pytest.mark.release)

        if filename in ACTIVE_FILES:
            item.add_marker(pytest.mark.active)
        elif filename in FROZEN_FILES:
            item.add_marker(pytest.mark.frozen)
        else:
            unknown.add(filename)

        if _matches_smoke(item):
            item.add_marker(pytest.mark.smoke)

    if unknown:
        pytest.exit(
            "MS-TEST1: unclassified test file(s): "
            + ", ".join(sorted(unknown))
            + ". Add each new test module to ACTIVE_FILES or FROZEN_FILES in tests/conftest.py.",
            returncode=4,
        )
