from __future__ import annotations

from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[2]
ISS = ROOT / "packaging" / "windows" / "ProgressStudio.iss"
BUILD = ROOT / "scripts" / "build-windows-installer.ps1"
CHECKLIST = ROOT / "packaging" / "windows" / "WIN3_CHECKLIST.md"


@pytest.mark.acceptance
@pytest.mark.release
def test_win3_installer_wraps_portable_payload_without_core_changes() -> None:
    text = ISS.read_text(encoding="utf-8")
    assert 'Source: "..\\..\\dist\\ProgressStudio\\*"' in text
    assert 'DefaultDirName={autopf}\\Progress Studio' in text
    assert "PrivilegesRequired=lowest" in text
    assert 'Filename: "{app}\\{#MyAppExeName}"' in text
    assert "[Uninstall" not in text  # use Inno's standard uninstall behavior


@pytest.mark.acceptance
@pytest.mark.release
def test_win3_installer_has_shortcuts_and_optional_desktop_task() -> None:
    text = ISS.read_text(encoding="utf-8")
    assert 'Name: "{group}\\Progress Studio"' in text
    assert 'Name: "{autodesktop}\\Progress Studio"' in text
    assert 'Tasks: desktopicon' in text
    assert 'Name: "desktopicon"' in text


@pytest.mark.acceptance
@pytest.mark.release
def test_win3_build_revalidates_known_good_portable_payload() -> None:
    text = BUILD.read_text(encoding="utf-8")
    assert "validate-windows-portable.ps1" in text
    assert "-PortableFolder $PortableFolder" in text
    assert "Inno Setup 6" in text
    assert "Get-FileHash" in text
    assert "SHA256" in text


@pytest.mark.acceptance
@pytest.mark.release
def test_win3_checklist_covers_install_uninstall_and_scope_boundary() -> None:
    text = CHECKLIST.read_text(encoding="utf-8")
    for phrase in (
        "Install acceptance",
        "Uninstall acceptance",
        "Start Menu",
        "Desktop shortcut",
        "Create Progress",
        "Mapping",
        "Payment",
        "Rebuild",
        "code signing",
        "licensing or activation",
        "SHA-256",
    ):
        assert phrase in text


@pytest.mark.acceptance
@pytest.mark.release
def test_win31_powershell_validator_does_not_use_stale_last_exit_code() -> None:
    text = BUILD.read_text(encoding="utf-8")
    validator_call = '& "$RepoRoot\\scripts\\validate-windows-portable.ps1" -PortableFolder $PortableFolder'
    assert validator_call in text
    tail = text.split(validator_call, 1)[1].split("# The .iss", 1)[0]
    assert "$LASTEXITCODE" not in tail
    assert 'if (-not $?)' in tail


@pytest.mark.acceptance
@pytest.mark.release
def test_windows_installer_build_guide_records_repeatable_process() -> None:
    guide = ROOT / "packaging" / "windows" / "BUILD_INSTALLER.md"
    text = guide.read_text(encoding="utf-8")
    for phrase in (
        "build-windows-portable.ps1",
        "validate-windows-portable.ps1",
        "build-windows-installer.ps1",
        "Inno Setup 6",
        "ISCC.exe",
        "ProgressStudio-Setup-2.3.0.exe",
        "SHA-256",
    ):
        assert phrase in text
