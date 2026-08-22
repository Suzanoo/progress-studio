from openpyxl import Workbook

from progress_studio.infrastructure.excel.worksheet_filters import configure_filter_buttons


def test_main_filter_ui_only_exposes_row_type_and_pa_buttons():
    wb = Workbook()
    ws = wb.active
    ws.append(["Row Type", "WBS", "Description", "P/A", "Activity ID", "Amount"])
    ws.append(["Activity", "1.1", "Test", "P", "A1000", 100])

    configure_filter_buttons(
        ws,
        header_row=1,
        last_row=2,
        last_col=6,
        visible_columns={1, 4},
    )

    assert ws.auto_filter.ref == "A1:F2"
    hidden = {fc.colId for fc in ws.auto_filter.filterColumn if fc.showButton is False}
    assert hidden == {1, 2, 4, 5}
