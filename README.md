# Progress Studio V2

Progress Studio converts a Primavera P6 XML export into an Excel workbook for schedule, amount, progress, plan distribution, and OKD reporting.

**Release:** 2.1.3

## Pipeline

```text
Primavera XML
→ ImportStep
→ ScheduleStep
→ TimescaleStep
→ AmountStep
→ ProgressStep
→ DistributionStep
→ OkdStep
```

The application does not use subprocesses or the former scripts `01` through `07`.

## Requirements

- Python 3.10 or newer
- Microsoft Excel is recommended for recalculating formulas before uploading to OKD

Install the dependency:

```bash
pip install -r requirements.txt
```

## Run interactively

```bash
python main.py
```

The application opens an XML file picker, asks for the weekly cutoff day, and asks for the plan distribution method when required.

## Run with command-line options

```bash
python main.py --input example/example.xml --cutoff-day 5 --amount 1000
```

Options:

- `--input`: Primavera P6 XML path
- `--cutoff-day`: `1` Monday through `7` Sunday
- `--amount`: placeholder amount used only when the XML does not contain activity amounts

Display help:

```bash
python main.py --help
```

## Amount rules

- When the XML contains activity amount data, Progress Studio uses those amounts.
- When XML amount mode is active and an activity has no amount, the activity receives `0`.
- When the XML contains no activity amount data, the placeholder amount is assigned to every activity.
- `Amount Mapping` remains editable.
- Project and parent WBS rows are displayed but are not directly mapped.

## Activity ID rules

- The Primavera Activity ID is retained when available.
- A missing Activity ID receives a deterministic generated ID such as `GEN-00015`.
- Activity rows are never exported with a blank Activity ID.

## Output

The project output includes the working Excel files and a final workbook containing:

- `main`
- `Amount Mapping`
- `Distribution Report`
- `progress`
- `progress_table`

The exact intermediate filenames are managed centrally by `progress_studio/config/settings.py`.

## OKD workflow

The final workbook contains live Excel formulas.

1. Open the generated workbook in Microsoft Excel.
2. Wait for formula calculation to finish.
3. Save the workbook.
4. Upload the saved workbook to OKD.

## Architecture

```text
progress_studio/
├── app/             # application, composition root, context, and pipeline runner
├── config/          # settings and workbook schema
├── domain/          # framework-independent data models
├── infrastructure/  # Primavera XML, filesystem, and Excel implementations
├── pipeline/        # seven executable pipeline steps
├── presentation/    # English CLI and distribution prompts
├── services/        # workflow and business services
└── version.py       # release version
main.py              # bootstrap entry point
```

Maintenance rules:

- `main.py` remains a small bootstrap file.
- Domain models do not depend on `openpyxl`.
- Excel details stay under `infrastructure/excel`.
- Workbook names and worksheet names are centralized under `config`.
- New features must not bypass the pipeline or reintroduce subprocess-based stages.

## Tests

Run the complete release test suite:

```bash
python -m unittest discover -s tests -v
```

Acceptance documents for MS-1 through MS-8 are included in the project root.


## Architecture cleanup

Version 2.0.1 removes the obsolete root-level `excel_toolkit/` and `distribution/` packages. Distribution rules now live under `progress_studio/services/distribution/`, Activity ID rules live in `progress_studio/domain/activity_id.py`, and Excel theme helpers live in `progress_studio/infrastructure/excel/styles.py`.

## Desktop application

Desktop Phase 2 adds a Tkinter workflow while retaining the original CLI.

Run the desktop application:

```bash
python desktop.py
```

Desktop Phase 2 includes:

- Primavera XML file picker
- weekly cutoff and fallback amount settings
- Auto, Flat, Front, Back, or Bell plan distribution
- background pipeline execution so the window remains responsive
- seven-step progress bar and live activity log
- automatic output under `Desktop/Progress_Studio_Output/`
- buttons to open the completed workbook and project folder

The desktop layer calls the same OOP services and pipeline as the CLI. Business and Excel logic are not duplicated inside Tkinter.

### Windows quick start

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python desktop.py
```

The CLI remains available:

```powershell
python main.py
```


## Documentation

- `README_ROADMAP.md` — current V3 milestone roadmap.
- `docs/milestones/` — implementation notes for completed milestones.
- `docs/acceptance/` — acceptance criteria for completed milestones.
- `CHANGELOG.md` — release and hotfix history.
- `COPILOT.md` — engineering rules for coding agents.
