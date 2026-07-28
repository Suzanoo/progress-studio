from pathlib import Path


def desktop_path() -> Path:
    home = Path.home()
    candidates = (
        home / "Desktop",
        home / "OneDrive" / "Desktop",
        home / "OneDrive - Personal" / "Desktop",
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    desktop = home / "Desktop"
    desktop.mkdir(parents=True, exist_ok=True)
    return desktop
