# WIN-2 Clean-Machine Validation

WIN-2 validates the **already-built WIN-1 portable folder** from the point of
view of a normal Windows user. It does not rebuild the application and does not
change Progress Studio core behavior.

## Gate A — isolated probe on the build machine

Run from the repository root. The portable folder may be anywhere; it does not
need to live under the repository.

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\validate-windows-portable.ps1 `
  -PortableFolder "C:\path\to\dist\ProgressStudio"
```

The script:

1. copies the portable folder to a fresh `%TEMP%` location outside the repo;
2. rejects common development-only directories inside the portable bundle;
3. clears Python/venv/Conda environment variables and reduces `PATH` to Windows
   system directories;
4. launches `ProgressStudio.exe --win1-smoke` from the isolated copy;
5. records bundle size, file count, and executable SHA-256;
6. removes the temporary portable copy unless `-KeepCopy` is supplied.

**Pass condition:** `WIN-2 AUTOMATED ISOLATION PROBE PASS`.

## Gate B — clean Windows user / VM

Copy the complete `ProgressStudio` folder to a Windows account or VM that does
not have the Progress Studio source repository, `.venv`, or build environment.
Python should not be required.

Do not copy only `ProgressStudio.exe`; WIN-1/WIN-2 are one-folder builds and the
`_internal` directory is part of the application.

### Launch and resource acceptance

- Double-click `ProgressStudio.exe`.
- No console window is required.
- Home/Welcome opens.
- Theme, icons, and workspace navigation are present.
- Closing and reopening the app succeeds.

### Create Progress acceptance

Use a known-good MSP XML input, then repeat with a known-good P6 XML input.
Verify:

- Create Progress completes without Python/runtime errors.
- `main`, `main_monthly`, and `Dashboard` are created.
- X/W/M reporting labels follow the frozen RN contract.
- Dashboard/overlay charts open without Excel repair/recovery prompts.
- workbook protection and user-editable controls behave as expected.

### Mapping / Payment / Rebuild acceptance

- Mapping workspace opens and can load its expected inputs.
- Payment workspace opens and its basic workflow is available.
- Rebuild workspace opens.
- Run at least one Progress rebuild and one Payment rebuild.
- If practical on the validation machine, exercise the full 2x2 matrix:
  Snapshot/Live × Progress/Payment.

### Excel acceptance

On a machine with Microsoft Excel installed:

- generated workbook opens without repair/recovery dialogs;
- F9/manual recalculation behavior follows the frozen workbook contract;
- save/reopen succeeds;
- charts, cutoff overlays, and transparent presentation remain intact.

On a machine without Excel, Progress Studio itself should still launch; Excel
visual/manual acceptance is recorded separately rather than treated as an app
startup failure.

## WIN-2 exit criteria

WIN-2 is complete only when both gates pass:

- automated isolated portable probe: PASS;
- manual clean-user/VM checklist: PASS.

Record the validation Windows version, Progress Studio commit/tag, bundle
SHA-256, and any Excel version used. Installer work starts only after this gate.
