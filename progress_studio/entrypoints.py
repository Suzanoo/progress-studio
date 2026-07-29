"""Stable application entry points used by scripts and packaged builds."""

from __future__ import annotations

from collections.abc import Sequence


def desktop_main() -> int:
    """Launch the Progress Studio desktop application."""
    from progress_studio.presentation.gui.app import ProgressStudioDesktopApp

    app = ProgressStudioDesktopApp()
    app.mainloop()
    return 0


def cli_main(argv: Sequence[str] | None = None) -> int:
    """Run the command-line workflow."""
    from progress_studio.app import build_application

    return build_application().run(argv)
