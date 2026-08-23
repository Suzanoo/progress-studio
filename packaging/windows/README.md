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

From the repository root (an activated development environment is optional):

```powershell
.\scripts\build-windows-portable.ps1
```

The script:

1. creates a disposable `.build-venv` isolated from the system/development Python;
2. installs the `build` and `dev` optional dependency groups there;
3. runs the smoke gate from that isolated environment;
4. builds `packaging/windows/ProgressStudio.spec`;
5. checks bundled JSON/icon resources;
6. launches the executable in non-UI `--win1-smoke` mode;
7. removes `.build-venv` after the build.

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

## WIN-2 clean-machine validation

After WIN-1 produces a passing portable folder, validate that exact folder
without rebuilding it:

```powershell
.\scripts\validate-windows-portable.ps1 `
  -PortableFolder "C:\path\to\dist\ProgressStudio"
```

This performs an automated source/venv-independent isolation probe. Then use
[`WIN2_CHECKLIST.md`](WIN2_CHECKLIST.md) on a clean Windows user or VM. WIN-2 is
not complete until both the automated probe and the manual clean-machine
acceptance pass.


## WIN-3 installer

After WIN-2 passes, build the Windows installer with Inno Setup 6:

```powershell
.\scripts\build-windows-installer.ps1 -PortableFolder "C:\path\to\known-good\ProgressStudio"
```

The installer wraps the validated portable payload and provides per-user installation, Start Menu shortcut, optional Desktop shortcut, and uninstall. See [`WIN3_CHECKLIST.md`](WIN3_CHECKLIST.md).
