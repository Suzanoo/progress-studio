from pathlib import Path


def test_focus_mapping_contract_is_present():
    root = Path(__file__).resolve().parents[1]
    source = (root / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    for required in (
        "def _toggle_sidebar", "def _set_sidebar_collapsed",
        "def _toggle_focus_mapping", "def _exit_focus_mapping",
        'self.bind("<F11>"', 'self.bind("<Escape>"',
        'label="Focus Mapping"',
    ):
        assert required in source
