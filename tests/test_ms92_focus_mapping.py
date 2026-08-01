from pathlib import Path


def test_focus_mapping_contract_is_present():
    root = Path(__file__).resolve().parents[1]
    source = (root / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    for required in (
        "def _toggle_sidebar",
        "def _set_sidebar_collapsed",
        "def _toggle_focus_mapping",
        "def _exit_focus_mapping",
        'self.bind("<F11>"',
        'self.bind("<Escape>"',
        'label="Toggle Workbook Generator"',
    ):
        assert required in source


def test_mapping_is_primary_by_default():
    from progress_studio.infrastructure.layout_preferences import LayoutPreferences

    assert LayoutPreferences().generator_collapsed is True


def test_layout_preferences_keep_focus_workspace_state(tmp_path):
    from progress_studio.infrastructure.layout_preferences import (
        LayoutPreferences,
        LayoutPreferencesRepository,
    )

    repository = LayoutPreferencesRepository(tmp_path / "layout.json")
    expected = LayoutPreferences(
        mapping_inputs_collapsed=True,
        generator_collapsed=True,
        sidebar_collapsed=True,
        mapping_sash=640,
    )
    repository.save(expected)
    assert repository.load() == expected
