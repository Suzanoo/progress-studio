# MS-7 Acceptance — Regression and Release

MS-7 passes only when all conditions below are true:

1. The complete V2 pipeline runs from Primavera XML to the final OKD workbook.
2. The example regression dataset produces the approved sheet structure and row counts.
3. Formula links, weekly columns, schedule rows, and OKD rows pass automated checks.
4. `python main.py --help` exits successfully and remains English-only.
5. Invalid input exits with a non-zero status and a clear error.
6. No root scripts `01` through `07`, legacy adapters, or subprocess calls remain.
7. Application source contains no Thai text.
8. README documents installation, interactive use, non-interactive use, outputs, and architecture.
9. The package exposes release version `2.0.1`.
10. All automated tests pass from a clean project directory.
