# Progress Studio 2.3.0 — User Guide

Progress Studio creates a progress workbook from schedule XML, lets you map BOQ amounts to activities, saves your mapping work, and exports the mapped workbook.

## Complete workflow

```text
Schedule XML
    ↓
Create Progress Workbook
    ↓
Load Progress Workbook + BOQ Workbook
    ↓
Map BOQ amounts to activities
    ↓
Save mapping session
    ↓
Export mapped workbook
    ↓
Open in Microsoft Excel, recalculate, and save
```

## Start here

- [Quick Start](QUICK_START.md)
- [Schedule XML Requirements](XML_REQUIREMENTS.md)
- [BOQ Workbook Requirements](BOQ_REQUIREMENTS.md)
- [BOQ Mapping Guide](MAPPING_GUIDE.md)
- [Sessions and Export](SESSIONS_EXPORT.md)
- [Troubleshooting](TROUBLESHOOTING.md)

## Important rules

- The original XML and BOQ files are read-only.
- Schedule XML import stops if any activity is missing Activity Name, Plan Start, or Plan Finish.
- Activity ID is optional; Progress Studio generates one when it is missing.
- WBS is optional; Progress Studio creates a flat structure when WBS data is unavailable.
- Export updates the `main` worksheet and writes mapping records. Formula worksheets recalculate when the exported workbook is opened in Microsoft Excel.
