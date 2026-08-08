from __future__ import annotations

from openpyxl.workbook.workbook import Workbook


# Excel calculation engine version written by current Microsoft 365 builds.
# A non-zero calcId prevents Excel from treating every generated workbook as if
# it came from an unknown/ancient calculation engine and rebuilding all formula
# caches on every open.  Existing non-zero IDs are preserved.
CURRENT_EXCEL_CALC_ID = 191029


def configure_incremental_excel_recalculation(workbook: Workbook) -> None:
    """Use Excel's dependency-based automatic calculation without full rebuilds.

    openpyxl writes formulas but does not evaluate them. Excel will calculate
    formula cells in Automatic mode as needed, while ``fullCalcOnLoad`` and
    ``forceFullCalc`` would unnecessarily force the whole workbook (tens of
    thousands of cells in real projects) to recalculate every time it opens.
    """

    calculation = workbook.calculation
    calculation.calcMode = "auto"
    calculation.fullCalcOnLoad = False
    calculation.forceFullCalc = False
    calculation.calcOnSave = True
    if not calculation.calcId:
        calculation.calcId = CURRENT_EXCEL_CALC_ID


def request_full_excel_recalculation(workbook: Workbook) -> None:
    """Explicit escape hatch for repair/debug workflows that need a full rebuild."""

    calculation = workbook.calculation
    calculation.calcMode = "auto"
    calculation.fullCalcOnLoad = True
    calculation.forceFullCalc = True
    calculation.calcOnSave = True
    calculation.calcId = 0
