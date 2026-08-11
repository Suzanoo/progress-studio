
from __future__ import annotations

from openpyxl.utils import get_column_letter

from progress_studio.domain.main_dataset import MainDataset


def _rows(ws, dataset: MainDataset) -> dict[str, int]:
    headers = {name: col for name, col in dataset.headers}
    row_type_col = headers.get("row type")
    pa_col = headers.get("p/a")
    if row_type_col is None or pa_col is None:
        return {}

    found: dict[str, int] = {}
    for row in range(dataset.header_row + 1, ws.max_row + 1):
        row_type = str(ws.cell(row, row_type_col).value or "").strip().lower()
        pa = str(ws.cell(row, pa_col).value or "").strip().upper()
        if row_type == "project summary":
            if pa == "P":
                found["project_plan"] = row
            elif pa == "A":
                found["project_actual"] = row
        elif row_type == "s-curve":
            if pa == "P":
                found["plan"] = row
            elif pa == "AP":
                found["acc_plan"] = row
            elif pa == "A":
                found["actual"] = row
            elif pa == "AA":
                found["acc_actual"] = row
    return found


def apply_weekly_scurve_cutoff_contract(
    workbook,
    dataset: MainDataset,
    *,
    sheet_name: str = "main",
    cutoff_ref: str = "Dashboard!$K$5",
) -> bool:
    """Make weekly S-Curve rows authoritative and cutoff-aware.

    Plan and Acc.Plan always render the full baseline.
    Actual and Acc.Actual are blank after the selected Dashboard cutoff.
    Acc.Actual plateaus through gaps up to cutoff instead of extending beyond it.
    """
    if sheet_name not in workbook.sheetnames or not dataset.periods:
        return False

    ws = workbook[sheet_name]
    rows = _rows(ws, dataset)
    required = {"project_plan", "project_actual", "plan", "acc_plan", "actual", "acc_actual"}
    if not required.issubset(rows):
        return False

    first_col = dataset.periods[0].column
    first_letter = get_column_letter(first_col)
    header_row = dataset.header_row

    for period in dataset.periods:
        col = period.column
        letter = get_column_letter(col)
        header_cell = f"{letter}${header_row}"

        ws.cell(rows["plan"], col).value = f"={letter}{rows['project_plan']}"
        ws.cell(rows["acc_plan"], col).value = (
            f'=IF(COUNT(${first_letter}{rows["plan"]}:{letter}{rows["plan"]})=0,"",'
            f'SUM(${first_letter}{rows["plan"]}:{letter}{rows["plan"]}))'
        )
        ws.cell(rows["actual"], col).value = (
            f'=IF(OR({header_cell}>{cutoff_ref},COUNT({letter}{rows["project_actual"]})=0),'
            f'"",{letter}{rows["project_actual"]})'
        )
        ws.cell(rows["acc_actual"], col).value = (
            f'=IF({header_cell}>{cutoff_ref},"",'
            f'IF(COUNT(${first_letter}{rows["actual"]}:{letter}{rows["actual"]})=0,"",'
            f'SUM(${first_letter}{rows["actual"]}:{letter}{rows["actual"]})))'
        )

    return True
