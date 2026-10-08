<p align="center"><img src="progress_studio/assets/brand/progress_studio_icon.png" width="88" alt="Progress Studio S-curve icon"></p>

# Progress Studio

**Plan in P6 or MSP. Report progress in Excel.**

Progress Studio connects your schedule XML to a practical weekly/monthly Excel
progress workbook. Keep your scheduling tool and familiar Excel reporting workflow;
stop rebuilding the same S-curve workbook by hand.

## Why it exists

Planning teams often maintain schedules in Primavera P6 or Microsoft Project,
then recreate progress tables and S-curves in Excel for reporting. Progress Studio
bridges that gap. It is a focused desktop reporting tool, not a replacement
scheduling system or an accounting application.

## Product flow

1. Export **P6 XML** or **MSP XML**.
2. **Create** an Excel workbook: choose Weight basis, Weekly cutoff and Plan distribution.
3. Open Excel, review the plan/S-curve and enter progress in editable cells in `main`.
4. **F9 and Save** recalculate formulas. Use **Rebuild** when generated views or
   structural inputs need refreshing; it starts from the saved workbook, not XML.

**BOQ Mapping** and **Payment** are optional, not mandatory setup steps.
**Earned Value** is available through Rebuild.

## What you can do

- Import P6 and MSP XML without changing source scheduling data.
- Choose **Equal**, **Duration**, or an explicitly selected numeric **Amount** field.
- Generate weekly/monthly Plan and Actual reporting, Dashboard and S-curve views.
- Allocate BOQ amounts to activities with **Mapping**.
- Prepare persistent **Payment Input** and generate **Payment Breakdown**.
- Rebuild **Progress** or **Payment** using their existing **Snapshot / Live** workflows.
- Generate live **Earned Value** using **Activity Amount** or **BOQ Mapping** as the
  monetary source; retain PV/EV, WBS performance and source-specific detail.
- Use an **ENG / THA** desktop interface and a short in-app **Help / Quick Guide**.
  Language changes apply on the next application launch.

Live does not mean every change needs only F9: structural BAC/source changes
require EV refresh, and Python-generated snapshots require their owning rebuild.
Equal/Duration weights are not money. See [EV monetary sources](docs/EV_MONETARY_SOURCES.md).

## Product status

Accepted source milestones:

- V1: `d0db778` (`progress-studio-v1-stable`).
- V1.1: `b16ddb0` (`progress-studio-v1.1-stable`).
- Weekly Compact V1: `c9087ea` (`weekly-compact-v1-stable`).
- Weekly Compact Final Presentation: `17cc3f3` (`weekly-compact-final-presentation-stable`).

The Final Presentation was accepted in Desktop Excel after F9, edit/recalculation,
cutoff movement, WBS outline, and Save/Close/Reopen checks. Windows full regression:
839 passed, 6 skipped, 6 subtests passed, 0 failed. This is source/workbook
acceptance, not a claim that a new installer or distribution was released.
Mini Gantt is **cancelled**, not a planned feature.
Finance implementation is retained for compatibility but hidden from normal V1.1
navigation. No Finance workflow is advertised here.

## Quick start / Windows

For an existing authorized portable build, extract the whole folder and run
`ProgressStudio.exe` (do not copy the EXE alone).
For an installer, follow the installation wizard.

To build from source, use the step-by-step
[Windows build guide](packaging/windows/WINDOWS_BUILD_GUIDE.md).
Release artifacts, signing and Desktop acceptance must be verified by the repository
owner; the documentation does not imply a new downloadable release has been published.

## Documentation

- [User workflow](docs/USER_WORKFLOW.md)
- [V1.1 desktop experience and acceptance](docs/V1_1_PRODUCT_EXPERIENCE.md)
- [Weight selection](docs/WEIGHT_SELECTION.md)
- [Earned Value monetary sources](docs/EV_MONETARY_SOURCES.md)
- [Payment Breakdown](docs/PAYMENT_BREAKDOWN.md)
- [Weekly Compact presentation](docs/WEEKLY_COMPACT.md)
- [Open next-step questions](docs/NEXT_STEPS.md)
- [Documentation index](docs/README.md)
- [Changelog](CHANGELOG.md) / [Roadmap and status](docs/ROADMAP.md)

## Development / tests

Requires Python 3.10+ with Tkinter. Existing compatibility entry points remain:
`desktop.py`, `main.py`, `progress-studio` and `progress-studio-cli`.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python desktop.py
python -m pytest
```

See [Development](docs/DEVELOPMENT.md), [Testing](docs/TESTING.md),
[Architecture](docs/ARCHITECTURE.md) and [Release checklist](docs/RELEASE_CHECKLIST.md).
Workbook formulas, sheet names and source identities are not translated.
