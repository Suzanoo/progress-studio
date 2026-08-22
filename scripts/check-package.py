from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REQUIRED = {
    "progress_studio/config/dashboard_theme.json",
    "progress_studio/config/payment_lines.json",
    "progress_studio/config/theme.json",
    "progress_studio/services/distribution/distribution_rules.json",
    "progress_studio/assets/dashboard/icons/actual.png",
    "progress_studio/assets/dashboard/icons/planned.png",
    "progress_studio/assets/dashboard/icons/schedule.png",
    "progress_studio/assets/dashboard/icons/time_impact.png",
}


def build_wheel(root: Path, wheel_dir: Path) -> Path:
    cmd = [
        sys.executable, "-m", "pip", "wheel", str(root),
        "--no-deps", "--no-build-isolation", "--wheel-dir", str(wheel_dir),
    ]
    subprocess.run(cmd, check=True)
    wheels = sorted(wheel_dir.glob("progress_studio-*.whl"))
    if len(wheels) != 1:
        raise RuntimeError(f"Expected one progress-studio wheel, found {len(wheels)}")
    return wheels[0]


def verify_wheel(wheel: Path) -> None:
    with zipfile.ZipFile(wheel) as zf:
        names = set(zf.namelist())
    missing = sorted(REQUIRED - names)
    if missing:
        raise RuntimeError("Wheel is missing required runtime resources: " + ", ".join(missing))


def smoke_import_from_wheel(wheel: Path, extract_dir: Path) -> None:
    with zipfile.ZipFile(wheel) as zf:
        zf.extractall(extract_dir)
    code = r"""
from pathlib import Path
from progress_studio import __version__
from progress_studio.entrypoints import cli_main, desktop_main
from progress_studio.presentation.gui.theme import PALETTE
from progress_studio.services.distribution.auto import load_rules
from progress_studio.config.payment_theme import CONFIG_PATH
import progress_studio.infrastructure.excel.dashboard_workbook as dashboard

assert __version__
assert callable(cli_main) and callable(desktop_main)
assert PALETTE.primary
assert load_rules()
assert CONFIG_PATH.is_file()
assert (Path(dashboard.__file__).resolve().parents[2] / "config" / "dashboard_theme.json").is_file()
for filename in ("planned.png", "actual.png", "schedule.png", "time_impact.png"):
    icon = Path(dashboard.__file__).resolve().parents[2] / "assets" / "dashboard" / "icons" / filename
    assert icon.is_file(), icon
"""
    env = dict(__import__("os").environ)
    env["PYTHONPATH"] = str(extract_dir)
    subprocess.run([sys.executable, "-c", code], check=True, cwd=extract_dir, env=env)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build and verify Progress Studio package resources.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    root = args.root.resolve()
    with tempfile.TemporaryDirectory(prefix="progress-studio-package-") as temp:
        temp_path = Path(temp)
        wheel = build_wheel(root, temp_path / "wheel")
        verify_wheel(wheel)
        smoke_import_from_wheel(wheel, temp_path / "installed")
        print(f"Package check passed: {wheel.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
