from __future__ import annotations

from datetime import date

from openpyxl import Workbook

from progress_studio.infrastructure.excel.dashboard_workbook import (
    DASHBOARD_SHEET,
    build_dashboard,
)


def _workbook() -> Workbook:
    wb = Workbook()
    progress = wb.active
    progress.title = "progress"
    progress.append(["project_start", "project_finish", "week_start", "plan", "actual"])
    progress.append([date(2026,1,1), date(2026,2,28), date(2026,1,2), 20, 10])
    progress.append([date(2026,1,1), date(2026,2,28), date(2026,1,9), 60, 25])

    table = wb.create_sheet("progress_table")
    table.append([
        "WBS", "Activities", "Amount", "P/A", "%Progress",
        date(2026,1,2), date(2026,1,9), "_Kind",
    ])
    table.append(["1.1", "Activity A", 1000, "P", 60, 20, 40, "activity"])
    table.append(["1.1", "Activity A", 1000, "A", 25, 10, 15, "activity"])
    table.append(["1.2", "Activity B", 500, "P", 100, 50, 50, "activity"])
    table.append(["1.2", "Activity B", 500, "A", 100, 50, 50, "activity"])
    return wb


def test_activity_progress_keeps_two_rows_and_adds_variance_status() -> None:
    wb = _workbook()
    build_dashboard(wb)
    ws = wb[DASHBOARD_SHEET]

    assert ws["N38"].value == "Variance"
    assert ws["P38"].value == "Status"

    assert ws["F39"].value == "Plan"
    assert ws["F40"].value == "Actual"
    assert ws["N39"].value == ""
    assert ws["P39"].value == ""
    assert ws["N40"].value == "=IFERROR(L40-L39,0)"
    assert '"Behind"' in ws["P40"].value
    assert '"Complete"' in ws["P40"].value

    # Outline grouping remains paired.
    assert ws.row_dimensions[39].outlineLevel == ws.row_dimensions[40].outlineLevel


def test_status_focus_dropdown_exists_without_column_filter_arrows() -> None:
    wb = _workbook()
    build_dashboard(wb)
    ws = wb[DASHBOARD_SHEET]

    assert ws["N37"].value == "Status"
    assert ws["P37"].value == "All"

    validations = list(ws.data_validations.dataValidation)
    status_validations = [
        v for v in validations
        if v.formula1 == '"All,Behind,On Track,Complete,Not Started"'
    ]
    assert len(status_validations) == 1

    # No AutoFilter dropdowns anywhere on Activity Progress.
    assert ws.auto_filter.ref is None

    # Macro-free focus rule dims whole Plan/Actual pairs based on Actual-row Status.
    rules = []
    for sqref in ws.conditional_formatting:
        if "B39:Q" in str(sqref):
            rules.extend(ws.conditional_formatting[sqref])
    assert rules
    formula_text = " ".join(
        formula
        for rule in rules
        for formula in (rule.formula or [])
    )
    assert "$P$37" in formula_text
    assert "INDEX($P:$P" in formula_text


def test_variance_status_columns_do_not_change_okd_progress_table_contract() -> None:
    wb = _workbook()
    original_headers = [cell.value for cell in wb["progress_table"][1]]
    build_dashboard(wb)

    assert [cell.value for cell in wb["progress_table"][1]] == original_headers
