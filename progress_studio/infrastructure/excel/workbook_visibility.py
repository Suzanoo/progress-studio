from __future__ import annotations


VISIBLE_SHEETS = (
    "README",
    "main",
    "main_monthly",
    "Payment Input",
    "Payment-Breakdown",
    "Payment",
    "Dashboard",
    "Earned Value",
    "EV Table",
)

HIDDEN_SHEETS = (
    "progress",
    "progress_table",
    "Dashboard_Data",
    "EV_Data",
)


def apply_final_sheet_visibility(
    workbook,
) -> tuple[tuple[str, ...], tuple[str, ...], tuple[str, ...]]:
    """Apply the final user-facing workbook visibility contract.

    Earned Value is an existing public feature and must stay visible whenever
    present. EV_Data remains a normal hidden helper sheet.
    """
    visible: list[str] = []
    hidden: list[str] = []
    very_hidden: list[str] = []

    for sheet in workbook.worksheets:
        if sheet.title in VISIBLE_SHEETS:
            sheet.sheet_state = "visible"
            visible.append(sheet.title)
        elif sheet.title in HIDDEN_SHEETS:
            sheet.sheet_state = "hidden"
            hidden.append(sheet.title)
        else:
            sheet.sheet_state = "veryHidden"
            very_hidden.append(sheet.title)

    if not visible and workbook.worksheets:
        workbook.worksheets[0].sheet_state = "visible"
        visible.append(workbook.worksheets[0].title)
        if workbook.worksheets[0].title in hidden:
            hidden.remove(workbook.worksheets[0].title)
        if workbook.worksheets[0].title in very_hidden:
            very_hidden.remove(workbook.worksheets[0].title)

    return tuple(visible), tuple(hidden), tuple(very_hidden)
