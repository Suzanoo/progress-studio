from openpyxl import Workbook

from progress_studio.infrastructure.excel.workbook_visibility import (
    apply_final_sheet_visibility,
)


def test_pb4r_final_visibility_keeps_earned_value_public():
    wb = Workbook()
    wb.active.title = "main"
    for name in (
        "Payment-Breakdown",
        "Dashboard",
        "Earned Value",
        "EV Table",
        "EV_Data",
        "BOQ Activity Mapping",
    ):
        wb.create_sheet(name)

    visible, hidden, very_hidden = apply_final_sheet_visibility(wb)

    assert "Payment-Breakdown" in visible
    assert "Earned Value" in visible
    assert "EV Table" in visible
    assert wb["Earned Value"].sheet_state == "visible"
    assert wb["EV Table"].sheet_state == "visible"

    assert "EV_Data" in hidden
    assert wb["EV_Data"].sheet_state == "hidden"

    assert "BOQ Activity Mapping" in very_hidden
    assert wb["BOQ Activity Mapping"].sheet_state == "veryHidden"
