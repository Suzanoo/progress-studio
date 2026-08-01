from pathlib import Path


def source() -> str:
    root = Path(__file__).resolve().parents[1]
    return (root / "progress_studio/presentation/gui/amount_mapping.py").read_text(encoding="utf-8")


def test_command_toolbar_contract():
    text = source()
    for label in ("Open Progress", "Open BOQ", "Save Session", "Undo", "Map", "Unmap", "Export"):
        assert f'text="{label}"' in text


def test_keyboard_first_contract():
    text = source()
    for sequence in ("<Control-o>", "<Command-o>", "<Control-s>", "<Command-s>", "<Control-z>", "<Command-z>", "<Delete>"):
        assert sequence in text


def test_loading_empty_state_and_notifications_exist():
    text = source()
    assert "loading_overlay" in text
    assert "activity_empty_var" in text
    assert "boq_empty_var" in text
    assert "def _notify" in text
    assert "mapping_status_bar" in text
