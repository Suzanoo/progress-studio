# Build the Windows Installer

This is the repeatable Windows release procedure for Progress Studio 2.3.0.

The installer is a wrapper around a **known-good portable build**. Do not build
an installer from an unvalidated portable folder.

## 1. Prerequisites

On the Windows build machine:

- Python 3.11 or another supported Python used by the project;
- Inno Setup 6;
- the Progress Studio source repository;
- a portable `ProgressStudio` folder that has already passed WIN-1 and WIN-2.

PowerShell may block unsigned local scripts. For the current PowerShell window
only, use:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

This setting disappears when that PowerShell process is closed.

## 2. Build the portable application when needed

From the repository root:

```powershell
.\scripts\build-windows-portable.ps1
```

Expected output:

```text
dist\
└─ ProgressStudio\
   ├─ ProgressStudio.exe
   └─ _internal\
```

The build script uses a disposable `.build-venv`; it must not require packages
to be installed into the user's global Python environment.

## 3. Validate the portable payload

Before creating an installer:

```powershell
.\scripts\validate-windows-portable.ps1 `
  -PortableFolder "C:\REAL\PATH\TO\ProgressStudio"
```

**Important:** `C:\REAL\PATH\TO\ProgressStudio` is an example. Replace it with
the actual folder on the build machine.

A passing payload reports:

```text
WIN-2 AUTOMATED ISOLATION PROBE PASS
```

Record its SHA-256. The installer must wrap that same known-good payload.

## 4. Locate Inno Setup when auto-detection does not find it

The compiler executable is `ISCC.exe`.

To locate it:

```powershell
Get-ChildItem `
  "C:\Program Files", `
  "C:\Program Files (x86)", `
  "$env:LOCALAPPDATA" `
  -Filter ISCC.exe -Recurse -ErrorAction SilentlyContinue |
Select-Object -ExpandProperty FullName
```

Copy the returned full path for the next step.

## 5. Build the installer

If Inno Setup is found automatically:

```powershell
.\scripts\build-windows-installer.ps1 `
  -PortableFolder "C:\REAL\PATH\TO\ProgressStudio"
```

If auto-detection does not find `ISCC.exe`, pass it explicitly:

```powershell
.\scripts\build-windows-installer.ps1 `
  -PortableFolder "C:\REAL\PATH\TO\ProgressStudio" `
  -InnoCompiler "C:\REAL\PATH\TO\ISCC.exe"
```

The script:

1. re-runs the WIN-2 isolation probe;
2. stages the validated portable payload;
3. invokes Inno Setup;
4. writes the installer under `dist\installer`;
5. prints the installer SHA-256.

Expected output:

```text
dist\
└─ installer\
   └─ ProgressStudio-Setup-2.3.0.exe
```

## 6. Install / uninstall acceptance

Run the installer and verify:

- Progress Studio appears in Windows Installed Apps;
- Start Menu launch works;
- optional Desktop shortcut works if selected;
- Home/Welcome, themes, and icons are present;
- Create Progress, Mapping, Payment, and Rebuild workspaces open;
- a known-good generated workbook opens in Excel without a repair prompt;
- uninstall is available and removes application files/shortcuts;
- user-created XML/Excel/project files outside the install directory remain untouched.

See [`WIN3_CHECKLIST.md`](WIN3_CHECKLIST.md) for the full checklist.

## 7. Current known-good WIN-3 artifact

Validated on Windows 10 build 19045 on 23-Aug-2026.

```text
Installer:
ProgressStudio-Setup-2.3.0.exe

Installer SHA-256:
76728AF64EB4D772351B4985D7AD612E9689C1474479EF54824714025EAEC73A

Portable payload SHA-256:
96F5D56348A8692659E3FFA0ED7238CE0E1632E4DB8E67AB25B6B09572A211A0
```

The installer binary itself is a release artifact and should not be committed to
the Git repository. Git stores the packaging source, validation record, and tag.
