from __future__ import annotations


VISIBLE_SHEETS = (
    "main",
    "main_monthly",
    "Payment Input",
    "Payment",
    "Dashboard",
)


def apply_final_sheet_visibility(workbook) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """Apply the final portable-workbook visibility contract.

    Visible user sheets:
      main, main_monthly, Payment Input, Payment, Dashboard

    Every other existing sheet becomes ``veryHidden`` so it stays available to
    Progress Studio / Excel formulas while remaining out of normal Unhide menus.
    Missing visible sheets are simply ignored because not every workflow creates
    Payment/Payment Input yet.
    """
    visible = []
    very_hidden = []

    for sheet in workbook.worksheets:
        if sheet.title in VISIBLE_SHEETS:
            sheet.sheet_state = "visible"
            visible.append(sheet.title)
        else:
            sheet.sheet_state = "veryHidden"
            very_hidden.append(sheet.title)

    # Excel requires at least one visible worksheet. A valid Progress Studio
    # workbook always has main, but keep the policy fail-safe deterministic.
    if not visible and workbook.worksheets:
        workbook.worksheets[0].sheet_state = "visible"
        visible.append(workbook.worksheets[0].title)
        if workbook.worksheets[0].title in very_hidden:
            very_hidden.remove(workbook.worksheets[0].title)

    return tuple(visible), tuple(very_hidden)
