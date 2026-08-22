# Progress Studio Pre-Production Roadmap

The current goal is not to add more features. The goal is to make the existing product understandable, editable, testable, performant and package-ready for Windows and macOS.

## Current position

Completed stabilization foundations include:

- MSP/P6 XML normalization and validation.
- Create Progress weekly/monthly/dashboard generation.
- reporting-range separation from display margins.
- `X / Wn / Mn` reporting-period labeling.
- Rebuild Snapshot/Live Progress and Payment ownership paths.
- workbook protection / visibility / F9-Save policy work.
- chart and overlay regression fixes.

The repository is now in cleanup and pre-production hardening.

## Milestones

| Milestone | Goal | Exit gate |
|---|---|---|
| **P0 Baseline Freeze** | Freeze a known-good product baseline and real reference workbooks | tagged checkpoint + clean Git state |
| **P1 Repository & Documentation Cleanup** | Clean docs/root/tests organization without changing product behavior | smoke + documentation/link checks |
| **P2 Architecture Contract Freeze** | Keep one current architecture/ownership source of truth | architecture review + cross-workspace ownership tests |
| **P3 Editable Workbook Rebuild** | Prove Rebuild after user adds/removes Activities, WBS and timescale structure | mutation matrix passes |
| **P4 Workflow Regression Matrix** | End-to-end Create / Mapping / Payment / Rebuild verification | Create + Rebuild 2x2 + representative workflows pass |
| **P5 Performance & Memory** | Reduce unnecessary openpyxl I/O and benchmark realistic projects | agreed load/save/runtime/RAM targets pass |
| **P6 Excel Compatibility** | Validate OOXML, formulas, charts, protection and recalculation in real Excel | no Excel repair dialog; compatibility gate passes |
| **P7 Desktop Production UX** | Final user workflow, errors, logs, Welcome and configuration polish | manual user acceptance |
| **P8 Windows Production Build** | Create Windows distributable/installer and test a clean machine | Windows 10/11 clean-machine gate |
| **P9 macOS Production Build** | Create macOS bundle and signing/notarization plan | supported Mac clean-machine gate |
| **P10 Release Candidate** | Feature freeze and full release regression | RC tag + full acceptance |
| **P11 Production Release** | Publish versioned production artifacts and documentation | signed/tagged release |

## P3 mutation matrix (production blocker)

Rebuild must be tested against structural workbook edits, including:

- add Activity;
- delete Activity;
- add WBS;
- delete WBS;
- reorder/move Activity under another WBS;
- extend timescale earlier/later;
- shrink timescale;
- insert/delete timescale periods.

The intent is that `main` remains a usable editable source workbook rather than a read-only report.

## P4 workflow matrix

At minimum:

```text
Create Progress
  MSP XML
  P6 XML

Rebuild
                    Progress    Payment
Snapshot               ✓           ✓
Live                   ✓           ✓
```

Mapping and Payment-only paths must preserve data owned by other workspaces.

## Release discipline

- No architecture refactor solely to fix a local bug.
- Product behavior changes require regression coverage.
- Full release tests are required before RC/production, not after every small edit.
- Real project workbooks should remain part of manual acceptance alongside synthetic fixtures.
