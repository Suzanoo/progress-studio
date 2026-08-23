"""PyInstaller-only launcher for the Windows portable build.

This module intentionally lives outside ``progress_studio`` so packaging behavior
can evolve without changing the frozen application/core entry points.
"""

from __future__ import annotations

import json
import sys
from importlib import resources


def _resource_smoke() -> None:
    """Fail fast when essential runtime resources are missing from the bundle."""
    import openpyxl  # noqa: F401
    from PIL import Image  # noqa: F401

    from progress_studio.entrypoints import desktop_main  # noqa: F401
    from progress_studio.presentation.gui.app import ProgressStudioDesktopApp  # noqa: F401

    required = (
        "config/theme.json",
        "config/dashboard_theme.json",
        "config/payment_lines.json",
        "services/distribution/distribution_rules.json",
        "assets/dashboard/icons/planned.png",
        "assets/dashboard/icons/actual.png",
        "assets/dashboard/icons/schedule.png",
        "assets/dashboard/icons/time_impact.png",
    )
    root = resources.files("progress_studio")
    for relative in required:
        ref = root.joinpath(*relative.split("/"))
        if not ref.is_file():
            raise FileNotFoundError(f"Missing packaged resource: progress_studio/{relative}")
        if relative.endswith(".json"):
            json.loads(ref.read_text(encoding="utf-8"))


def main() -> int:
    if "--win1-smoke" in sys.argv:
        _resource_smoke()
        return 0
    if "--version" in sys.argv:
        from progress_studio.version import __version__

        # Windowed binaries do not have a console, but this remains useful when
        # launched from cmd/PowerShell with console redirection during QA.
        print(__version__)
        return 0

    from progress_studio.entrypoints import desktop_main

    return desktop_main()


if __name__ == "__main__":
    raise SystemExit(main())
