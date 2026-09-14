# Progress Studio — Windows Build Guide

This guide is the operator-facing, end-to-end procedure for producing and validating a Windows release build of Progress Studio.

The packaging pipeline is:

```text
Source repository
    ↓
WIN-1 — PyInstaller one-folder portable build
    ↓
WIN-2 — isolated portable validation
    ↓
WIN-3 — Inno Setup installer
    ↓
Manual install / uninstall / Excel acceptance
```

Use the existing technical acceptance checklists for deeper validation details:

- `packaging/windows/WIN2_CHECKLIST.md`
- `packaging/windows/WIN3_CHECKLIST.md`

Do not bypass a failed gate for a release build.

---

## 1. Prerequisites

Build on Windows.

Required:

- Git
- Python compatible with the project (`>=3.10`)
- the repository cloned locally
- Inno Setup 6 for WIN-3 installer packaging
- Microsoft Excel for the final workbook visual/manual acceptance

Recommended:

- build from the current accepted `main`
- start from a clean working tree
- keep the normal development `.venv` available for local verification; the WIN-1 script creates its own isolated build environment

The project build dependencies are declared in `pyproject.toml`, including PyInstaller under the `build` optional dependency.

---

## 2. Open PowerShell at the repository

Example:

```powershell
cd C:\Users\Suzanoo\Dev\progress-studio
```

Confirm the repository state:

```powershell
git switch main
git pull

git status
git log -3 --oneline --decorate
```

For a release build, `git status` should normally end with:

```text
nothing to commit, working tree clean
```

If the working tree contains intentional changes, finish and verify them before building the release artifact.

---

## 3. Activate the development environment

If the repository `.venv` already exists:

```powershell
.\.venv\Scripts\Activate.ps1
```

Check Python:

```powershell
python --version
```

The packaging scripts do not depend on the development virtual environment being embedded into the application. WIN-1 creates a separate isolated build environment.

---

## 4. Allow repository PowerShell scripts for this session

Run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

This changes policy only for the current PowerShell process.

---

# WIN-1 — Build the portable application

## 5. Build the PyInstaller one-folder bundle

From the repository root:

```powershell
.\scripts\build-windows-portable.ps1
```

The script performs the release-oriented portable build flow, including the pre-build smoke gate and isolated build environment.

Expected successful ending:

```text
WIN-1 PASS
```

Expected output location:

```text
dist\ProgressStudio\
dist\ProgressStudio\ProgressStudio.exe
```

### Important: this is a one-folder build

Do not distribute only `ProgressStudio.exe`.

The complete `dist\ProgressStudio` folder is the application payload. Files under `_internal` and the other packaged resources are required.

### If the pre-build smoke gate fails

Stop the release build and fix the failing test or regression first.

Do not use `-SkipTests` merely to get a release artifact past a known failure.

After fixing the issue, rerun:

```powershell
.\scripts\build-windows-portable.ps1
```

---

# WIN-2 — Validate the portable build

## 6. Run the isolated portable probe

Run:

```powershell
.\scripts\validate-windows-portable.ps1 `
  -PortableFolder "$PWD\dist\ProgressStudio"
```

Expected successful ending:

```text
WIN-2 AUTOMATED ISOLATION PROBE PASS
```

WIN-2 copies the built application outside the repository, removes development Python/venv/Conda influence from the launch environment, and runs the packaged application smoke probe.

The command also reports:

- file count
- bundle size
- executable SHA-256
- validation report path

Do not continue to WIN-3 if WIN-2 fails.

## 7. Clean-user / VM acceptance

For a release candidate, also perform the manual Gate B procedure in:

```text
packaging/windows/WIN2_CHECKLIST.md
```

Use a Windows account or VM without the Progress Studio source repository or development environment.

At minimum verify:

- `ProgressStudio.exe` launches without Python installed
- Home / Welcome loads correctly
- theme and icons are present
- Create Progress opens and can generate a known-good workbook
- Mapping opens
- Payment opens
- Rebuild opens
- generated workbook opens in Excel without repair/recovery prompts
- save/reopen keeps workbook charts and overlays intact

---

# WIN-3 — Build the Windows installer

## 8. Verify Inno Setup 6

The installer build requires Inno Setup 6 and its command-line compiler `ISCC.exe`.

First check the common installation path:

```powershell
Test-Path "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```

Expected result when installed there:

```text
True
```

Another common per-user path is:

```text
%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe
```

`where.exe ISCC.exe` may return nothing even when Inno Setup is installed because `ISCC.exe` is not necessarily added to `PATH`.

To search common locations:

```powershell
Get-ChildItem `
  "C:\Program Files", `
  "C:\Program Files (x86)", `
  "$env:LOCALAPPDATA\Programs" `
  -Filter ISCC.exe `
  -Recurse `
  -ErrorAction SilentlyContinue |
  Select-Object -ExpandProperty FullName
```

If Inno Setup is not installed and `winget` is available:

```powershell
winget install --id JRSoftware.InnoSetup -e
```

## 9. Build the installer

When `ISCC.exe` is located at:

```text
C:\Program Files (x86)\Inno Setup 6\ISCC.exe
```

run:

```powershell
.\scripts\build-windows-installer.ps1 `
  -PortableFolder "$PWD\dist\ProgressStudio" `
  -InnoCompiler "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```

The WIN-3 script re-runs the WIN-2 isolated validation by default before creating the installer.

For a normal release build, do not pass:

```text
-SkipPortableProbe
```

Expected successful ending:

```text
WIN-3 BUILD PASS
```

Expected installer:

```text
dist\installer\ProgressStudio-Setup-2.3.0.exe
```

The script also prints the installer SHA-256.

---

# Manual installer acceptance

## 10. Install the generated setup package

Run:

```text
dist\installer\ProgressStudio-Setup-2.3.0.exe
```

Verify:

- installation succeeds without Python or `.venv`
- default per-user installation does not request administrator elevation
- Start Menu shortcut launches Progress Studio
- optional Desktop shortcut launches Progress Studio
- Home / Welcome loads
- theme and icons load
- Create Progress opens
- Mapping opens
- Payment opens
- Rebuild opens
- one known-good Create Progress input produces a workbook
- generated workbook opens normally in Excel
- Progress Studio closes and reopens normally

Follow the full checklist in:

```text
packaging/windows/WIN3_CHECKLIST.md
```

---

## 11. Uninstall acceptance

Use Windows **Installed apps** or the Progress Studio uninstall entry.

Verify:

- uninstall completes normally
- installed application files are removed
- Start Menu / Desktop shortcuts are removed
- user-created XML, Excel, and project files outside the install directory remain untouched

---

## 12. Same-version reinstall acceptance

After uninstalling, install the same generated setup package again.

Verify Progress Studio launches normally again.

---

# Short operator sequence

For an already prepared Windows build machine, the normal release sequence is:

```powershell
cd C:\Users\Suzanoo\Dev\progress-studio

git switch main
git pull

git status
git log -3 --oneline --decorate

.\.venv\Scripts\Activate.ps1

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

.\scripts\build-windows-portable.ps1

.\scripts\validate-windows-portable.ps1 `
  -PortableFolder "$PWD\dist\ProgressStudio"

.\scripts\build-windows-installer.ps1 `
  -PortableFolder "$PWD\dist\ProgressStudio" `
  -InnoCompiler "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```

Then perform the manual WIN-3 install/uninstall/reinstall acceptance.

---

# Release record

For every accepted Windows release build, record at least:

- Progress Studio commit SHA / tag
- Windows version used for validation
- Excel version used for workbook acceptance
- portable validation SHA-256/report
- installer SHA-256
- date of manual acceptance

This allows an installer to be traced back to the exact source and validation environment that produced it.

---

# Current packaging boundaries

The current Windows packaging milestone deliberately does not provide:

- code signing
- auto-update
- licensing / activation
- online accounts or billing
- feature gating

Those concerns belong to later release/distribution milestones and should not be mixed into the current WIN-1/WIN-2/WIN-3 packaging contract.
