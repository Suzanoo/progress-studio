# MS-1 — Standard Application Entry Points

## Delivered

- Added `pyproject.toml` as the package and dependency source of truth.
- Added `progress-studio` desktop command.
- Added `progress-studio-cli` command-line command.
- Added `python -m progress_studio` desktop launch support.
- Added stable `desktop_main()` and `cli_main()` functions.
- Preserved `desktop.py` and `main.py` as compatibility launchers.
- Added smoke tests for version and entry-point imports.
- Documented Windows and macOS editable installation.

## Not included

Executable packaging, installers, application merging, and cross-application workflow integration are reserved for later milestones.
