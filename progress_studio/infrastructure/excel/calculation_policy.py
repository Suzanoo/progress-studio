from __future__ import annotations

from openpyxl.workbook.workbook import Workbook


def request_full_excel_recalculation(workbook: Workbook) -> None:
    """Make Microsoft Excel rebuild every formula result when the file opens.

    openpyxl writes formulas but does not calculate them. Setting calcId to zero
    prevents Excel from trusting cached results created by another calculation
    engine/version.
    """

    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcId = 0
