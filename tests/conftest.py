from __future__ import annotations

from pathlib import Path

import pytest


# Product test profiles are behavior-based, not milestone-based.
# Every test belongs to exactly one primary profile by directory:
#   tests/unit/         - isolated logic / parsers / domain helpers
#   tests/integration/  - component and workflow boundaries
#   tests/regression/   - bugs/product contracts that must never return
#   tests/acceptance/   - release-level product acceptance contracts
# Every collected test also belongs to `release`.
PROFILE_DIRS = {
    "unit": "unit",
    "integration": "integration",
    "regression": "regression",
    "acceptance": "acceptance",
}

# Smoke is intentionally small and high-value. It is a subset of the profiles
# above and should answer: "is this build safe enough to hand to a user?"
SMOKE_NODEIDS = {
    "integration/desktop/test_entrypoints.py",
    "integration/create_progress/test_workbook_generation_engine.py",
    "integration/create_progress/test_snapshot_progress_table.py",
    "integration/rebuild/test_rebuild_core_contract.py",
    "integration/rebuild/test_progress_rebuild_engine.py::test_rb2_rebuild_progress_replaces_only_progress_owned_sheets",
    "integration/rebuild/test_snapshot_performance.py",
    "integration/rebuild/test_payment_only_rebuild.py::test_rb4_rebuild_payment_replaces_payment_only",
    "integration/rebuild/test_payment_collision_lanes.py",
    "integration/payment/test_payment_position_engine.py",
    "integration/payment/test_payment_line_renderer.py::test_ms_pay65_default_render_includes_every_populated_payment",
    "regression/dashboard/test_dashboard_reporting_range.py",
    "regression/dashboard/test_chart_ooxml_integrity.py",
    "regression/overlay/test_traditional_overlay.py",
    "regression/workbook/test_final_workbook_policy.py",
}


def _relative_nodeid(item: pytest.Item) -> str:
    nodeid = item.nodeid.replace("\\", "/")
    return nodeid.split("tests/", 1)[-1]


def _primary_profile(item: pytest.Item) -> str | None:
    path = Path(str(item.fspath)).resolve()
    parts = path.parts
    try:
        tests_index = parts.index("tests")
    except ValueError:
        return None
    if tests_index + 1 >= len(parts):
        return None
    folder = parts[tests_index + 1]
    return PROFILE_DIRS.get(folder)


def _matches_smoke(item: pytest.Item) -> bool:
    short = _relative_nodeid(item)
    filename = Path(str(item.fspath)).name
    for target in SMOKE_NODEIDS:
        if "::" in target:
            if short == target:
                return True
        elif short == target or short.startswith(target + "::") or filename == target:
            return True
    return False


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    invalid: list[str] = []

    for item in items:
        item.add_marker(pytest.mark.release)
        profile = _primary_profile(item)
        if profile is None:
            invalid.append(_relative_nodeid(item))
        else:
            item.add_marker(getattr(pytest.mark, profile))

        if _matches_smoke(item):
            item.add_marker(pytest.mark.smoke)

    if invalid:
        pytest.exit(
            "Test organization: test file(s) outside unit/integration/regression/acceptance: "
            + ", ".join(sorted(invalid)),
            returncode=4,
        )
