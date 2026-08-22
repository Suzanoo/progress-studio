from __future__ import annotations

from tests._paths import REPO_ROOT, TESTS_ROOT


def test_pytest_product_profiles_are_registered() -> None:
    pyproject = (REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8")
    for marker in ("unit:", "smoke:", "integration:", "regression:", "acceptance:", "release:"):
        assert marker in pyproject
    assert 'addopts = "--strict-markers"' in pyproject
    assert "active:" not in pyproject
    assert "frozen:" not in pyproject


def test_product_profile_helper_scripts_exist() -> None:
    scripts = REPO_ROOT / "scripts"
    for name in (
        "test-unit.ps1",
        "test-smoke.ps1",
        "test-integration.ps1",
        "test-regression.ps1",
        "test-acceptance.ps1",
        "test-release.ps1",
    ):
        assert (scripts / name).is_file()
    assert not (scripts / "test-active.ps1").exists()
    assert not (scripts / "test-frozen.ps1").exists()


def test_tests_are_organized_by_product_profile_directory() -> None:
    allowed = {"unit", "integration", "regression", "acceptance", "fixtures"}
    top_level_tests = list(TESTS_ROOT.glob("test_*.py"))
    assert not top_level_tests
    for child in TESTS_ROOT.iterdir():
        if child.is_dir() and not child.name.startswith("__"):
            assert child.name in allowed


def test_conftest_no_longer_contains_milestone_tier_lists() -> None:
    source = (TESTS_ROOT / "conftest.py").read_text(encoding="utf-8")
    assert "FROZEN_FILES" not in source
    assert "ACTIVE_FILES" not in source
    assert "PROFILE_DIRS" in source
