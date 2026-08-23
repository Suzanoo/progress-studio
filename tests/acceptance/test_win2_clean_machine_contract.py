from __future__ import annotations

from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "scripts" / "validate-windows-portable.ps1"
CHECKLIST = ROOT / "packaging" / "windows" / "WIN2_CHECKLIST.md"


@pytest.mark.acceptance
@pytest.mark.release
def test_win2_validation_harness_is_repository_independent() -> None:
    text = SCRIPT.read_text(encoding="utf-8")

    assert "Copy-Item -Path $PortableFolder -Destination $RunRoot -Recurse" in text
    assert '"--win1-smoke"' in text
    assert 'Remove-Item "Env:$Name"' in text
    assert 'VIRTUAL_ENV' in text
    assert 'CONDA_PREFIX' in text
    assert '$env:PATH = "$env:SystemRoot\\System32;$env:SystemRoot"' in text
    assert "Start-Process" in text
    assert "python -m" not in text.lower()
    assert "-m pytest" not in text.lower()


@pytest.mark.acceptance
@pytest.mark.release
def test_win2_rejects_development_directories_in_portable_bundle() -> None:
    text = SCRIPT.read_text(encoding="utf-8")
    for name in (".git", ".venv", ".build-venv", ".pytest_cache", "tests", "build"):
        assert f'"{name}"' in text


@pytest.mark.acceptance
@pytest.mark.release
def test_win2_checklist_covers_real_user_workflows() -> None:
    text = CHECKLIST.read_text(encoding="utf-8")
    required = (
        "Create Progress",
        "MSP XML",
        "P6 XML",
        "Mapping",
        "Payment",
        "Rebuild",
        "Snapshot/Live × Progress/Payment",
        "Microsoft Excel",
        "F9",
        "SHA-256",
    )
    for phrase in required:
        assert phrase in text
