import pytest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.skip(reason="Frozen pre-LW command label contract: Export is now contextual as Export Mapped Workbook.")
def test_project_command_bar_contract():
    text = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    for label in ("Open", "Save", "Save As", "Undo", "Map", "Unmap", "Export"):
        assert f'("{label}",' in text or f'text="{label}"' in text


def test_user_facing_project_language_replaces_session_language():
    mapping = (ROOT / "progress_studio/presentation/gui/amount_mapping.py").read_text(encoding="utf-8")
    assert 'title="Save Progress Studio project"' in mapping
    assert '"Recent Projects"' in mapping
    assert 'text="Save Session"' not in mapping


def test_keyboard_first_contract():
    app = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    for key in ("<Control-o>", "<Control-s>", "<Control-Shift-S>", "<Control-z>"):
        assert key in app
