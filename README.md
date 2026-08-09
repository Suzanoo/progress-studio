
### MS-P1.7 — Dashboard from initial XML generation

- The `Dashboard` sheet is created during the **Build progress sheets** stage, before BOQ mapping.
- A user importing XML receives a dashboard-ready workbook immediately, even while all activity amounts are still zero.
- The same four-level orange Activity Data hierarchy theme is applied during initial generation and mapped export.
- Mapping export refreshes the existing dashboard; it is no longer the first point at which the dashboard appears.

# Progress Studio

Progress Studio creates a progress workbook from schedule XML, maps BOQ amounts to activities, saves mapping sessions, and exports a recalculation-ready Excel workbook.

**Current release:** `2.3.0`

## User documentation

- [Thai User Guide](docs/th/README.md)
- [English User Guide](docs/en/README.md)
- [Documentation index](docs/README.md)

Start with:

- [English User Guide](docs/en/README.md)
- [Quick Start](docs/en/QUICK_START.md)
- [Schedule XML Requirements](docs/en/XML_REQUIREMENTS.md)
- [BOQ Mapping Guide](docs/en/MAPPING_GUIDE.md)
- [Troubleshooting](docs/en/TROUBLESHOOTING.md)

## Schedule XML contract

Every activity must contain:

```text
Activity Name
Plan Start
Plan Finish
```

Import stops and no workbook is created when any required value is missing or invalid.

Optional:

- Activity ID — generated automatically when missing
- WBS — a flat structure is created when missing
- Calendar, relationships, duration, actual dates, progress, resources, and codes

## Windows quick start

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python desktop.py
```

## Workflow

```text
Schedule XML
→ Progress Workbook
→ Load BOQ
→ Map BOQ to Activities
→ Save Project (.progressstudio; self-contained in v8)
→ Rebuild Latest Workbook
→ Open in Microsoft Excel, recalculate, and save
```


## MS-R1 — Self-contained workbook rebuild

Projects saved by this version embed verified copies of the Progress and BOQ source workbooks inside the `.progressstudio` project. After the project has been saved once in v8, the original source files are no longer required to rebuild a workbook with the latest Progress Studio generation/export engine.

Legacy v7 and older projects still open normally, but require their original/relinked workbooks once. Save the project again to upgrade it to the self-contained v8 format. Migration of Actual Progress from a separately edited legacy workbook is a later milestone and is not part of MS-R1.

## Technical documentation

- [Documentation index](docs/README.md)
- [Architecture](ARCHITECTURE.md)
- [Roadmap](README_ROADMAP.md)
- [Changelog](CHANGELOG.md)
- [Release checklist](RELEASE_CHECKLIST.md)
- [Engineering rules](COPILOT.md)
- `docs/milestones/` — milestone implementation records
- `docs/acceptance/` — milestone acceptance records

## Tests

```powershell
python -m unittest discover -s tests -v
```

---

## Standard installation and entry points (MS-1)

Create a project-local virtual environment and install Progress Studio as an editable package.

### Windows PowerShell

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e ".[dev]"
progress-studio
```

### macOS / Linux

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -e ".[dev]"
progress-studio
```

Available entry points:

```text
progress-studio           Open the desktop GUI
python -m progress_studio Open the desktop GUI
progress-studio-cli       Run the command-line workflow
python desktop.py         Legacy-compatible desktop launcher
python main.py            Legacy-compatible CLI launcher
```

Run the test suite with:

```bash
pytest
```

## Excel export theme configuration

The exported `main` sheet uses one central theme configuration file:

```text
progress_studio/infrastructure/excel/export_theme.py
```

It contains two independent palettes:

- `TimescalePalette` — Project, WBS level 1, WBS level 2, Plan/Actual, and S-curve colors in the timescale section.
- `ActivityDataPalette` — WBS level 1, 2, 3, and 4 colors in the Activity Data section. WBS levels deeper than 4 reuse the level-4 color.

The Activity Data formatter is implemented in:

```text
progress_studio/infrastructure/excel/activity_data_theme.py
```

Activity Data styling changes fill and font only. It uses the `Outline Level` value, preserves existing row borders, and does not modify timescale cells. Levels 1–4 use progressively lighter fills; level 5 and deeper use the level-4 fill.

## Actual Amount calculation

The exported `main` sheet calculates earned Actual Amount automatically:

- Activity Actual Amount = Activity Plan Amount × Actual `% Complete`
- WBS Actual Amount = sum of descendant Activity Actual Amounts
- Project Actual Amount = sum of all Activity Actual Amounts

Weekly Plan/Actual roll-ups continue to use the full Plan Amount as their weight, so displaying earned Actual Amount does not change progress percentages.


## Weekly and Monthly Main Views

Generated workbooks keep `main` as the editable **weekly source of truth** and now add `main_monthly` as a calculated monthly view. The monthly worksheet reuses the same Activity Data columns, Plan/Actual row pairs, WBS outline grouping, colors, and Row Type / P/A filters. Its timescale contains one column per reporting month, dated with the last weekly cutoff available in that month.

Monthly progress is formula-driven from `main`: normal Plan/Actual rows sum the weekly increments belonging to the month, while cumulative S-Curve rows take the last weekly cumulative value in the month. Editing weekly Actual progress in `main` therefore updates `main_monthly` when Excel recalculates. `main_monthly` is presentation-only and should not be used as a second progress-entry source.

## Excel Dashboard

Generated workbooks now include a separate **Dashboard** worksheet as the first tab.
Project-level KPI/chart data stays live from `progress`; Activity Progress reads the latest value-only snapshot in `progress_table`. The snapshot is regenerated on Export/Rebuild and contains:

- Cutoff Date dropdown
- Weekly / Monthly dropdown
- Four KPI cards: Planned Progress, Actual Progress, Schedule Status, and Time Impact
- S-Curve chart with full baseline Plan and cutoff-limited Actual
- Activity Progress summary

`Dashboard_Data` is a hidden helper worksheet used by the chart and dropdown logic. KPI values always follow the selected cutoff date; changing Weekly/Monthly only changes the chart reporting view. Dashboard colors and chart layout are configurable in `progress_studio/config/dashboard_theme.json`. `main` remains the editable Plan/Actual master, while `progress_table` is deliberately a lightweight snapshot refreshed by Export/Rebuild rather than a live weekly mirror.

### Payment line theme config

Payment backbone colors and lightweight label sizing are configured in:

`progress_studio/config/payment_lines.json`

The renderer reads this config for all populated Payment periods. `colors` maps
`P01`, `P02`, ... to hex colors, while `label` controls badge width, height,
font size, corner radius, text color, and anchor offset. Payment line geometry
remains cell-border based.

### Embedded Payment workflow

Payment now uses one workbook from preparation through rebuild.

User-facing sheets:
- `Dashboard`
- `main`
- `Payment Input`
- `Payment`
- `main_monthly`
- `progress`

Generated/support sheets remain available to the engine but are hidden. In particular,
`progress_table` is a value-only snapshot and is rebuilt (not patched) whenever the
Payment workbook is rebuilt. `Payment` is also deleted/recreated from the current
`main` + `Payment Input`.

`Payment Input` is persistent user data. Existing percentages are reconciled by
Activity ID; new activities receive suggested values. Payment Date is no longer an
input because Planned Eligible Date is calculated from the latest required Activity point.

### MS-RB1 — Standalone rebuild core contract

The standalone rebuild path treats the selected workbook itself as the project source.

- `main` is always the schedule/progress source of truth.
- Progress rebuild owns only: `main_monthly`, `progress`, `progress_table`, `Dashboard_Data`, `Dashboard`.
- Payment rebuild owns only: `Payment`, and requires embedded `Payment Input`.
- `main`, `Payment Input`, internal metadata, and unknown user sheets are preserved by contract.
- The rebuild core has no runtime dependency on XML, BOQ files, `.progressstudio`,
  `.boqstudio`, mapping allocations, or the working tree.
- Workbook analysis is sparse: it reads workbook metadata plus `main` worksheet XML only,
  avoiding full openpyxl workbook loading before a rebuild is selected.

### MS-RB2 — Progress rebuild execution

`WorkbookRebuildEngine.rebuild_progress()` now executes the Progress rebuild contract.

Input:
- any Progress Studio `.xlsx` / `.xlsm` workbook with a valid `main` sheet.

Preserved:
- `main`
- `Payment Input`
- `Payment`
- internal metadata sheets
- unknown/user-created sheets

Deleted and rebuilt from the current `main`:
- `main_monthly`
- `progress`
- `progress_table`
- `Dashboard_Data`
- `Dashboard`

The rebuild writes atomically through a temporary workbook. `progress_table` remains a
value-only hidden snapshot, Dashboard_Data remains hidden, and Excel calculation policy
is normalized before the final replace. MS-RB2 intentionally leaves the current
formula-driven monthly/progress/dashboard behavior unchanged; snapshot/performance
hardening belongs to MS-RB3.

### MS-RB3 — Snapshot performance hardening

Standalone `Rebuild Progress` now breaks the live dependency chain from generated
views back to `main`.

RB3 snapshot contract:
- `main_monthly` = value-only snapshot
- `progress` = value-only snapshot derived from current Activity data
- `progress_table` = value-only snapshot
- `Dashboard_Data` may keep lightweight formulas for Dashboard view/cutoff controls,
  but does not link to `main`
- `Dashboard` remains interactive and reads generated support sheets
- final calculation policy is always `auto`, `fullCalcOnLoad=False`,
  `forceFullCalc=False`

Existing export/monthly builders keep their previous default behavior; snapshot mode
is enabled by the standalone rebuild engine only.

Real NKC2_R03 benchmark:
- RB2 generated-sheet formulas: 10,911
- RB3 generated-sheet formulas: 2,306
- no direct generated-sheet formula links to `main`

### MS-RB3.1 — Dashboard source contract fix

Dashboard now follows the same two-sheet contract used by the OKD app:

- S-Curve + KPI source: `progress`
- Activity table source: `progress_table`

`progress` stores cumulative Plan/Actual as 0..100 percent-points. Dashboard_Data
converts those values to Excel chart fractions explicitly, including values below 1%
(e.g. 0.31% -> 0.0031).

Monthly chart points are sampled from cumulative `progress`:
- Monthly Plan = last weekly cutoff in the month
- Monthly Actual = last populated weekly Actual in the month

No monthly SUM or multi-area LOOKUP formula is used, preventing the previous
`#VALUE!` chain in Monthly Actual.

### MS-RB4 — Payment-only rebuild

`WorkbookRebuildEngine.rebuild_payment()` now owns the standalone Payment rebuild path.

Input:
- one workbook containing `main` + embedded `Payment Input`

Rebuilt:
- `Payment` only

Preserved:
- `main`
- `Payment Input`
- `main_monthly`
- `progress`
- `progress_table`
- `Dashboard_Data`
- `Dashboard`
- internal metadata and user-created sheets

The Payment renderer resolves Planned Eligible positions from current `main`, reads the
current Payment requirements from `Payment Input`, removes stale `Payment`, and writes a
new Payment sheet. Progress-generated views are deliberately not reconciled or rebuilt.

### MS-RB5 — Payment collision lanes and label polish

Payment business logic remains unchanged: each period has a true Planned Eligible
boundary calculated from its sparse Activity requirements.

Rendering now separates only the visual geometry when multiple Payment periods would
occupy the same Excel boundary. Nearby lanes are allocated using configurable offsets,
while every horizontal branch still ends at the true Activity target. Shifted header
notes explicitly mark the visual offset as display-only.

Default Payment label style is now:
- width: 145 px
- height: 26 px
- font: 12
- rounded corner radius: 6 px

Collision and label settings live in `progress_studio/config/payment_lines.json`,
including `collision_max_offset` and `collision_row_step`.

### MS-RB6 — Standalone Rebuild workspace

Rebuild is now a first-class workspace in the desktop sidebar:

`Home → Create Progress Bar → Mapping → Payment → AI Helper → Export → Rebuild → Settings`

The Rebuild workspace accepts one Excel workbook and asks for exactly one mode:

- **Progress Workbook** — rebuilds `main_monthly`, `progress`, `progress_table`,
  `Dashboard_Data`, and `Dashboard` from `main`.
- **Payment** — rebuilds `Payment` only from `main + Payment Input`.

The UI never asks for `.progressstudio`, `.boqstudio`, XML, BOQ, or mapping-tree inputs.

Workspace ownership is also cleaned up:
- **Export** creates the first mapped workbook only; old Rebuild controls are removed.
- **Payment** prepares/reconciles `Payment Input` only.
- **Rebuild** owns every post-Excel regeneration action.

Rebuild execution runs on a worker thread and writes through the RB2/RB4 atomic engines.

