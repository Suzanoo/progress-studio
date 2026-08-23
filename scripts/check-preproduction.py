from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ACCEPTANCE_DIR = ROOT / "tests/acceptance"


def run(label: str, *args: str) -> None:
    print(f"\n=== {label} ===", flush=True)
    completed = subprocess.run(args, cwd=ROOT)
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)


def main() -> int:
    python = sys.executable

    # Run the release suite by behavior directory rather than one monolithic
    # pytest process. Coverage is equivalent to the full collected suite, but
    # this avoids cumulative slowdowns in legacy end-to-end acceptance tests.
    run("Unit", python, "-m", "pytest", "tests/unit", "-q")
    run("Regression", python, "-m", "pytest", "tests/regression", "-q")
    run("Integration", python, "-m", "pytest", "tests/integration", "-q")

    for acceptance_file in sorted(ACCEPTANCE_DIR.glob("test_*.py")):
        run(f"Acceptance: {acceptance_file.name}", python, "-m", "pytest", str(acceptance_file), "-q")

    run("Smoke", python, "-m", "pytest", "-m", "smoke", "-q")
    run("Package artifact", python, "scripts/check-package.py")

    print("\nPR-1 pre-production gate passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
