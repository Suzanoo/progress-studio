from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def test_preproduction_gate_files_exist() -> None:
    assert (ROOT / "scripts" / "check-preproduction.py").is_file()
    assert (ROOT / "scripts" / "check-preproduction.ps1").is_file()
    assert (ROOT / "docs" / "PR1_PRE_PRODUCTION_FREEZE.md").is_file()


def test_preproduction_gate_covers_every_primary_test_directory() -> None:
    source = (ROOT / "scripts" / "check-preproduction.py").read_text(encoding="utf-8")
    for path in ("tests/unit", "tests/regression", "tests/integration", "tests/acceptance"):
        assert path in source
    assert "scripts/check-package.py" in source
    assert '"smoke"' in source
