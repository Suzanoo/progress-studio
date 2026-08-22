from progress_studio import __version__
from progress_studio.entrypoints import cli_main, desktop_main


def test_package_version_matches_release() -> None:
    assert __version__ == "2.3.0"


def test_entrypoints_are_callable() -> None:
    assert callable(desktop_main)
    assert callable(cli_main)
