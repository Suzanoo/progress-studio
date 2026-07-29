# MS-1 Acceptance Criteria — Application Skeleton

MS-1 passes only when every item below is true.

- [x] The target package structure exists.
- [x] `main.py` is a bootstrap entry point only.
- [x] `ProgressStudioApplication`, `Pipeline`, and `PipelineContext` are separate.
- [x] Dependencies are assembled in `app/bootstrap.py`.
- [x] CLI text, source comments, and documentation are English.
- [x] Workbook sheet names are centralized in `config/workbook_schema.py`.
- [x] The validated V1 workflow remains available through one temporary adapter.
- [x] Automated MS-1 tests pass.
- [x] `python main.py --help` exits successfully.

## Deliberately outside MS-1

- Replacing scripts `01` through `07`.
- Moving business rules into domain and service objects.
- Removing subprocess calls from the temporary legacy workflow.
- Changing workbook formulas, formatting, or outputs.

These items belong to MS-2 through MS-7.
