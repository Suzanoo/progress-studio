from pathlib import Path


def test_ms91_shell_modules_exist():
    root = Path(__file__).resolve().parents[1]
    assert (root / "progress_studio/presentation/gui/theme.py").is_file()
    assert (root / "progress_studio/presentation/gui/strings.py").is_file()


def test_ui_text_is_centralized_and_english_first():
    from progress_studio.presentation.gui.strings import tr

    assert tr("mapping_workspace") == "Mapping Workspace"
    assert tr("app_name") == "Progress Studio"


def test_settings_uses_release_version():
    from progress_studio.config.settings import SETTINGS
    from progress_studio.version import __version__

    assert SETTINGS.version == __version__


def test_production_shell_contract_is_present():
    source = (
        Path(__file__).resolve().parents[1]
        / "progress_studio/presentation/gui/app.py"
    ).read_text(encoding="utf-8")
    for required in (
        "def _build_menu",
        "def _build_sidebar",
        "def _build_generator_panel",
        "def _build_workspace_panel",
        '"AI Helper"',
        '"6  Export"',
    ):
        assert required in source
