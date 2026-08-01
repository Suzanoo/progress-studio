import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_ms91_shell_modules_exist():
    assert (ROOT / "progress_studio/presentation/gui/theme.py").is_file()
    assert (ROOT / "progress_studio/presentation/gui/strings.py").is_file()
    assert (ROOT / "progress_studio/config/theme.json").is_file()


def test_ui_text_is_centralized_and_english_first():
    from progress_studio.presentation.gui.strings import tr
    assert tr("mapping_workspace") == "Mapping Workspace"
    assert tr("app_name") == "Progress Studio"


def test_settings_uses_release_version():
    from progress_studio.config.settings import SETTINGS
    from progress_studio.version import __version__
    assert SETTINGS.version == __version__


def test_style_b_shell_contract_is_present():
    source = (ROOT / "progress_studio/presentation/gui/app.py").read_text(encoding="utf-8")
    for required in (
        "def _build_menu", "def _build_sidebar", "def _build_generator_panel",
        "def _build_mapping_workspace", "def _show_workspace", '"AI Helper"',
        '"Import"', '"Mapping"', '"Export"', '"Settings"',
    ):
        assert required in source


def test_theme_is_external_and_has_semantic_colors():
    theme = json.loads((ROOT / "progress_studio/config/theme.json").read_text(encoding="utf-8"))
    for key in ("primary", "primary_hover", "primary_soft", "sidebar", "surface", "selection"):
        assert key in theme["colors"]
