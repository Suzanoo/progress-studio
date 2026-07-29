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
