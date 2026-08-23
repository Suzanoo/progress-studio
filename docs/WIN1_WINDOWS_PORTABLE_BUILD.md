# WIN-1 — Windows Portable Build

WIN-1 is the first platform-specific packaging milestone after PR-1. Its purpose
is to prove that the frozen Progress Studio application can run from a Windows
portable folder without activating a Python environment.

## Frozen core

WIN-1 does not change Progress, Mapping, Payment, Rebuild, workbook calculation,
reporting-period, or Normalizer behavior. Windows packaging wraps the PR-1
application through the stable `progress_studio.entrypoints.desktop_main` entry
point.

## Build format

WIN-1 uses PyInstaller **one-folder** mode:

```text
dist/
└── ProgressStudio/
    ├── ProgressStudio.exe
    └── _internal/
```

One-folder is intentional for the first Windows milestone because it has faster
startup than one-file builds and makes missing resources easier to diagnose.

## Build

On Windows:

```powershell
.\scripts\build-windows-portable.ps1
```

The build script runs the smoke gate, builds the application, verifies required
resources, and executes a non-UI bundle smoke through
`ProgressStudio.exe --win1-smoke`.

## WIN-1 acceptance

After the automated build passes, manually verify:

1. Double-click opens the desktop UI with no console window.
2. Theme and dashboard icons are present.
3. Schedule XML file selection works.
4. Create Progress can generate a workbook.
5. Mapping, Payment, and Rebuild workspaces open.
6. Closing the application exits cleanly.

WIN-1 is not the customer installer. WIN-2 validates the portable bundle on a
clean Windows machine; installer, signing, update, and licensing work remain out
of scope.
