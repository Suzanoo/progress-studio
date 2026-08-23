# WIN-3 Windows Installer Acceptance

WIN-3 wraps the already validated WIN-1/WIN-2 portable payload. The installer must
not change Progress Studio workbook behavior.

## Build

Install Inno Setup 6 on the Windows build machine, then from the repository root:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\build-windows-installer.ps1 `
  -PortableFolder "C:\path\to\known-good\ProgressStudio"
```

The script re-runs the WIN-2 isolation probe before packaging unless
`-SkipPortableProbe` is explicitly supplied.

## Install acceptance

Run `dist\installer\ProgressStudio-Setup-2.3.0.exe`.

Verify:

- install succeeds without Python or a `.venv`;
- default per-user install does not require administrator elevation;
- Start Menu shortcut opens Progress Studio;
- optional Desktop shortcut opens Progress Studio;
- Home/Welcome, theme and icons load;
- Create Progress, Mapping, Payment and Rebuild workspaces open;
- use one known-good Create Progress input and open the generated workbook in Excel;
- installed app closes and reopens normally.

## Uninstall acceptance

Use Windows Apps/Installed apps or the Progress Studio uninstall entry.

Verify:

- uninstall completes normally;
- installed application files and shortcuts are removed;
- user-created XML/Excel/project files outside the install directory are not removed.

## Reinstall / same-version acceptance

Install the same WIN-3 build again after uninstalling. Verify launch again. WIN-3
does not yet promise migration between different product versions; upgrade policy
belongs to a later release milestone.

## WIN-3 exclusions

WIN-3 deliberately does **not** implement:

- code signing;
- auto-update;
- licensing or activation;
- online accounts/payment;
- feature gating.

Record Windows version, installer SHA-256, Progress Studio commit/tag, and Excel
version used for the manual acceptance.
