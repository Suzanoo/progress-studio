from __future__ import annotations

from pathlib import Path


def test_ms_test1_pytest_markers_are_registered() -> None:
    pyproject = (Path(__file__).parents[1] / "pyproject.toml").read_text(encoding="utf-8")
    for marker in ("smoke:", "active:", "frozen:", "release:"):
        assert marker in pyproject
    assert 'addopts = "--strict-markers"' in pyproject


def test_ms_test1_helper_scripts_exist() -> None:
    scripts = Path(__file__).parents[1] / "scripts"
    for name in (
        "test-smoke.ps1",
        "test-active.ps1",
        "test-frozen.ps1",
        "test-release.ps1",
    ):
        assert (scripts / name).is_file()


def test_ms_test1_new_modules_cannot_be_silently_unclassified() -> None:
    source = (Path(__file__).parent / "conftest.py").read_text(encoding="utf-8")
    assert "FROZEN_FILES" in source
    assert "ACTIVE_FILES" in source
    assert "unclassified test file" in source
