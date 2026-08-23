# Windows Packaging

WIN-1 proves that the frozen PR-1 application can run as a portable Windows
desktop application without a Python environment.

## WIN-1 output

```text
dist/
└── ProgressStudio/
    ├── ProgressStudio.exe
    └── _internal/
```

WIN-1 deliberately uses a **one-folder** PyInstaller build. Installer creation,
code signing, shortcuts, update channels, licensing, and activation are later
Windows release milestones.

## Build on Windows

From an activated development environment at the repository root:

```powershell
.\scripts\build-windows-portable.ps1
```

The script:

1. installs/validates the `build` optional dependency group;
2. runs the smoke gate;
3. builds `packaging/windows/ProgressStudio.spec`;
4. checks bundled JSON/icon resources;
5. launches the executable in non-UI `--win1-smoke` mode.

To rebuild quickly after the smoke gate has already passed:

```powershell
.\scripts\build-windows-portable.ps1 -SkipTests
```

## Manual acceptance after build

Double-click `dist/ProgressStudio/ProgressStudio.exe` and verify:

- the Home workspace opens without a terminal window;
- theme/icons load correctly;
- Schedule XML file selection works;
- Create Progress can generate a workbook;
- Mapping, Payment, and Rebuild workspaces open;
- the application exits normally.

Do not distribute WIN-1 as the customer installer. WIN-2 will validate this
portable folder on a clean Windows machine before installer work begins.
