"""WIN-1 verification for a built one-folder Windows distribution."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


REQUIRED_RELATIVE_FILES = (
    Path("ProgressStudio.exe"),
    Path("_internal/progress_studio/config/theme.json"),
    Path("_internal/progress_studio/config/dashboard_theme.json"),
    Path("_internal/progress_studio/config/payment_lines.json"),
    Path("_internal/progress_studio/services/distribution/distribution_rules.json"),
    Path("_internal/progress_studio/assets/dashboard/icons/planned.png"),
    Path("_internal/progress_studio/assets/dashboard/icons/actual.png"),
    Path("_internal/progress_studio/assets/dashboard/icons/schedule.png"),
    Path("_internal/progress_studio/assets/dashboard/icons/time_impact.png"),
)


def verify(folder: Path, *, run_smoke: bool = True) -> None:
    folder = folder.resolve()
    missing = [rel for rel in REQUIRED_RELATIVE_FILES if not (folder / rel).is_file()]
    if missing:
        rendered = "\n".join(f"  - {item}" for item in missing)
        raise SystemExit(f"WIN-1 portable bundle is missing required files:\n{rendered}")

    if run_smoke:
        exe = folder / "ProgressStudio.exe"
        result = subprocess.run(
            [str(exe), "--win1-smoke"],
            cwd=folder,
            timeout=60,
            check=False,
        )
        if result.returncode != 0:
            raise SystemExit(f"WIN-1 executable smoke failed with exit code {result.returncode}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", type=Path)
    parser.add_argument("--no-run", action="store_true", help="Only inspect bundle contents.")
    args = parser.parse_args(argv)
    verify(args.folder, run_smoke=not args.no_run)
    print(f"WIN-1 portable verification passed: {args.folder.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
