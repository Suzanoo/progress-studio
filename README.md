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
→ Save Session
→ Export mapped workbook
→ Open in Microsoft Excel, recalculate, and save
```

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
