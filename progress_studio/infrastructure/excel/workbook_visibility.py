from __future__ import annotations


VISIBLE_SHEETS = (
    "README",
    "main",
    "main_monthly",
    "Payment Input",
    "Payment",
    "Dashboard",
)


def apply_final_sheet_visibility(workbook) -> tuple[tuple[str, ...], tuple[str, ...], tuple[str, ...]]:
    """Apply the final portable-workbook visibility contract.

    Visible:
      README, main, main_monthly, Payment Input, Payment, Dashboard

    Hidden:
      every support/internal sheet, including progress, progress_table,
      Dashboard_Data, Info, Timescale Info, Amount Mapping and reports.

    Progress Studio intentionally avoids ``veryHidden`` in the working product
    build. Advanced users and project support can unhide helper sheets in Excel
    for inspection without modifying the workbook package. Protection still
    prevents accidental formula edits.
    """
    visible: list[str] = []
    hidden: list[str] = []

    for sheet in workbook.worksheets:
        if sheet.title in VISIBLE_SHEETS:
            sheet.sheet_state = "visible"
            visible.append(sheet.title)
        else:
            sheet.sheet_state = "hidden"
            hidden.append(sheet.title)

    # Excel requires at least one visible sheet.
    if not visible and workbook.worksheets:
        workbook.worksheets[0].sheet_state = "visible"
        visible.append(workbook.worksheets[0].title)
        if workbook.worksheets[0].title in hidden:
            hidden.remove(workbook.worksheets[0].title)

    return tuple(visible), tuple(hidden), ()
