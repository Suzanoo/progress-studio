from __future__ import annotations

import runpy
from pathlib import Path

import tomllib


ROOT = Path(__file__).resolve().parents[2]


def test_win1_build_files_exist() -> None:
    required = (
        ROOT / "packaging/windows/ProgressStudio.spec",
        ROOT / "packaging/windows/launcher.py",
        ROOT / "packaging/windows/verify_portable.py",
        ROOT / "scripts/build-windows-portable.ps1",
    )
    assert all(path.is_file() for path in required)


def test_pyproject_declares_pyinstaller_as_build_only_dependency() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert data["project"]["optional-dependencies"]["build"] == ["pyinstaller>=6.10,<7"]
    assert all("pyinstaller" not in dep.lower() for dep in data["project"]["dependencies"])


def test_win1_spec_collects_runtime_resources_without_changing_core_entrypoint() -> None:
    text = (ROOT / "packaging/windows/ProgressStudio.spec").read_text(encoding="utf-8")
    assert '"config/*.json"' in text
    assert '"assets/dashboard/icons/*.png"' in text
    assert '"services/distribution/*.json"' in text
    assert 'name="ProgressStudio"' in text
    assert "console=False" in text

    launcher = (ROOT / "packaging/windows/launcher.py").read_text(encoding="utf-8")
    assert "from progress_studio.entrypoints import desktop_main" in launcher
    assert "--win1-smoke" in launcher


def test_win1_launcher_resource_smoke_runs_from_source_tree(monkeypatch) -> None:
    launcher_path = ROOT / "packaging/windows/launcher.py"
    monkeypatch.setattr("sys.argv", [str(launcher_path), "--win1-smoke"])
    try:
        runpy.run_path(str(launcher_path), run_name="__main__")
    except SystemExit as exc:
        assert exc.code == 0


def test_win1_build_uses_disposable_isolated_environment() -> None:
    text = (ROOT / "scripts/build-windows-portable.ps1").read_text(encoding="utf-8")
    assert '".build-venv"' in text
    assert "-m venv $BuildVenv" in text
    assert '-e ".[build,dev]"' in text
    assert "& $BuildPython -m pytest -m smoke -q" in text
    assert "& $BuildPython -m PyInstaller" in text
    assert "Removing isolated build environment" in text

    gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")
    assert ".build-venv/" in gitignore
