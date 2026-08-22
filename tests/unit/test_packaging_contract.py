from __future__ import annotations

from pathlib import Path
import re
import tomllib

from progress_studio.version import __version__

ROOT = Path(__file__).resolve().parents[2]


def test_pyproject_packages_all_runtime_resources() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    patterns = set(data["tool"]["setuptools"]["package-data"]["progress_studio"])
    assert {
        "config/*.json",
        "assets/dashboard/icons/*.png",
        "services/distribution/*.json",
    } <= patterns


def test_declared_runtime_resources_exist_in_source_tree() -> None:
    expected = [
        ROOT / "progress_studio/config/dashboard_theme.json",
        ROOT / "progress_studio/config/payment_lines.json",
        ROOT / "progress_studio/config/theme.json",
        ROOT / "progress_studio/services/distribution/distribution_rules.json",
        ROOT / "progress_studio/assets/dashboard/icons/actual.png",
        ROOT / "progress_studio/assets/dashboard/icons/planned.png",
        ROOT / "progress_studio/assets/dashboard/icons/schedule.png",
        ROOT / "progress_studio/assets/dashboard/icons/time_impact.png",
    ]
    assert all(path.is_file() for path in expected)


def test_pyproject_and_runtime_version_match() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert data["project"]["version"] == __version__


def test_stable_entrypoints_are_declared() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    scripts = data["project"]["scripts"]
    assert scripts["progress-studio"] == "progress_studio.entrypoints:desktop_main"
    assert scripts["progress-studio-cli"] == "progress_studio.entrypoints:cli_main"


def test_root_launchers_delegate_to_stable_entrypoints() -> None:
    desktop = (ROOT / "desktop.py").read_text(encoding="utf-8")
    cli = (ROOT / "main.py").read_text(encoding="utf-8")
    assert "from progress_studio.entrypoints import desktop_main" in desktop
    assert "from progress_studio.entrypoints import cli_main" in cli


def test_gitignore_excludes_build_and_python_artifacts() -> None:
    text = (ROOT / ".gitignore").read_text(encoding="utf-8")
    required = [".venv/", ".pytest_cache/", "__pycache__/", "build/", "dist/", "*.egg-info/", "*.whl"]
    assert all(item in text for item in required)
