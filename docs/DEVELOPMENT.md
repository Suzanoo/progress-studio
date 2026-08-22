# Progress Studio Development Guide

## Setup

Windows PowerShell:

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e ".[dev]"
```

macOS/Linux:

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -e ".[dev]"
```

Run the desktop app:

```text
progress-studio
```

Run the CLI:

```text
progress-studio-cli
```

## Development rules

- Keep business/domain state out of GUI widgets.
- Activity identity is Activity ID.
- BOQ identity must use stable keys/source metadata, not Description alone.
- Preserve workbook ownership boundaries documented in `ARCHITECTURE.md`.
- A renderer owns only the workbook objects it creates.
- A local bug fix should not create a new parallel engine.
- Every behavior change requires a focused regression test.
- Use Git history for historical milestone details instead of growing active README files.

## Excel / openpyxl performance policy

openpyxl load/save work is expensive. Prefer:

```text
load/create once
-> derive in RAM
-> mutate the active workbook
-> finalize
-> save once
```

Do not reopen a workbook just to apply protection, visibility or calculation properties when the current workbook object is available.

Test/debug validation may reopen a saved file to inspect OOXML or persisted workbook properties.

## Repository hygiene

Do not commit:

- `.venv/`;
- `.pytest_cache/`;
- `__pycache__/` / `*.pyc`;
- `build/`, `dist/`, `*.egg-info/`;
- generated output workbooks, logs or temporary ZIPs unless they are intentional fixtures.

Historical design/milestone records belong under `docs/history/` and are not active product contracts.


## Package-content verification

Before platform-specific Windows/macOS packaging, build a wheel from the current source and verify that runtime JSON configuration and dashboard icons are present:

```powershell
.\scripts\check-package.ps1
```

The check builds into a temporary directory and leaves no `dist/` artifact in the repository. Platform-specific installer/bundle definitions belong under `packaging/`.

## Before handoff

At minimum run the focused tests for the changed subsystem plus Smoke:

```powershell
python -m pytest -m smoke -q
```

Before merge/tag/release, use the release gate documented in `docs/TESTING.md`.
