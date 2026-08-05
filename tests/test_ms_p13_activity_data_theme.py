from datetime import date

from openpyxl import Workbook

from progress_studio.infrastructure.excel.activity_data_theme import (
    apply_activity_data_wbs_hierarchy,
)


def _sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = "main"
    ws.append([])
    ws.append([])
    ws.append([])
    ws.append(["Row Type", "WBS", "Description", "P/A", "Outline Level", "Amount", date(2026, 1, 2)])
    ws.append(["Project Summary", "", "Project", "P", 0, 100, None])
    ws.append(["WBS", "1", "Parent", "P", 1, 100, None])
    ws.append([None, None, None, "A", None, None, None])
    ws.append(["WBS", "1.1", "Child", "P", 2, 100, None])
    ws.append([None, None, None, "A", None, None, None])
    ws.append(["Activity", "1.1.1", "Task", "P", 3, 100, None])
    return ws


def test_activity_data_hierarchy_colors_only_data_section():
    ws = _sheet()
    project_fill = ws.cell(5, 1).fill.fgColor.rgb
    activity_fill = ws.cell(10, 1).fill.fgColor.rgb
    timescale_fill = ws.cell(6, 7).fill.fgColor.rgb
    level_1_border = tuple(side.style for side in (ws.cell(6, 1).border.left, ws.cell(6, 1).border.right, ws.cell(6, 1).border.top, ws.cell(6, 1).border.bottom))
    level_2_border = tuple(side.style for side in (ws.cell(8, 1).border.left, ws.cell(8, 1).border.right, ws.cell(8, 1).border.top, ws.cell(8, 1).border.bottom))

    apply_activity_data_wbs_hierarchy(ws)

    assert ws.cell(6, 1).fill.fgColor.rgb.endswith("F4B183")
    assert ws.cell(7, 1).fill.fgColor.rgb.endswith("F4B183")
    assert ws.cell(8, 1).fill.fgColor.rgb.endswith("F8CBAD")
    assert ws.cell(9, 1).fill.fgColor.rgb.endswith("F8CBAD")
    assert tuple(side.style for side in (ws.cell(6, 1).border.left, ws.cell(6, 1).border.right, ws.cell(6, 1).border.top, ws.cell(6, 1).border.bottom)) == level_1_border
    assert tuple(side.style for side in (ws.cell(8, 1).border.left, ws.cell(8, 1).border.right, ws.cell(8, 1).border.top, ws.cell(8, 1).border.bottom)) == level_2_border
    assert ws.cell(5, 1).fill.fgColor.rgb == project_fill
    assert ws.cell(10, 1).fill.fgColor.rgb == activity_fill
    assert ws.cell(6, 7).fill.fgColor.rgb == timescale_fill
