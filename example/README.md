# Example Inputs

This directory contains small/reference inputs used for development and manual checks.

- `example.xml` — historical example schedule input.
- `example_BOQ.xlsx` — example BOQ workbook.
- `golden/progress.xlsx` — reference progress workbook used by a current integration regression test.

The old numbered-script walkthrough (`01_import_xml.py` ... `05_build_progress_workbook.py`) was removed because those scripts are no longer the current application workflow.

Use the installed application entry points instead:

```text
progress-studio
progress-studio-cli
```

See the repository [README](../README.md) and [User Workflow](../docs/USER_WORKFLOW.md).
