from __future__ import annotations

import os
import sys
from pathlib import Path


def user_data_dir() -> Path:
    """Return the per-user writable application-data directory."""
    if sys.platform == "win32":
        root = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
        return root / "ProgressStudio"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "ProgressStudio"
    root = Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))
    return root / "progress-studio"
