# MS-8 Acceptance Criteria — Architecture Cleanup

MS-8 passes only when all criteria below are satisfied:

1. The root-level `excel_toolkit/` package is removed.
2. The root-level `distribution/` package is removed.
3. The root-level `excel_theme.py` module is removed.
4. Activity ID rules live in `progress_studio/domain/activity_id.py`.
5. Distribution algorithms and automatic rules live in `progress_studio/services/distribution/`.
6. Excel styles and theme helpers live in `progress_studio/infrastructure/excel/styles.py`.
7. No application import references removed modules.
8. No circular imports exist inside `progress_studio`.
9. The end-to-end regression manifest remains unchanged.
10. The complete automated test suite passes.
