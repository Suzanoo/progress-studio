from __future__ import annotations

from calendar import monthrange
from datetime import date, datetime
import json
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

DASHBOARD_SHEET = "Dashboard"
DATA_SHEET = "Dashboard_Data"
PROGRESS_SHEET = "progress"
TABLE_SHEET = "progress_table"

NAVY = "17365D"
BLUE = "2F75B5"
GREEN = "70AD47"
RED = "C00000"
AMBER = "BF9000"
LIGHT_BLUE = "EAF2F8"
LIGHT_GREEN = "EAF4E3"
LIGHT_RED = "FCE8E6"
LIGHT_AMBER = "FFF4D6"
LIGHT_GRAY = "F3F4F6"
BORDER = "D9DEE7"
TEXT = "1F2937"
MUTED = "667085"
WHITE = "FFFFFF"


def _load_dashboard_theme() -> dict:
    path = Path(__file__).resolve().parents[2] / "config" / "dashboard_theme.json"
    defaults = {
        "font": "Aptos",
        "colors": {
            "navy": NAVY, "blue": BLUE, "green": GREEN, "red": RED,
            "amber": AMBER, "light_blue": LIGHT_BLUE, "light_green": LIGHT_GREEN,
            "light_red": LIGHT_RED, "light_amber": LIGHT_AMBER,
            "light_gray": LIGHT_GRAY, "border": BORDER, "text": TEXT,
            "muted": MUTED, "white": WHITE,
        },
        "layout": {"title": "PROGRESS STUDIO DASHBOARD", "activity_rows": 8, "chart_height": 8.0, "chart_width": 20.5},
    }
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return defaults
    result = defaults.copy()
    result["colors"] = {**defaults["colors"], **loaded.get("colors", {})}
    result["layout"] = {**defaults["layout"], **loaded.get("layout", {})}
    result["font"] = loaded.get("font", defaults["font"])
    return result


_THEME = _load_dashboard_theme()
_FONT = _THEME["font"]
_COLORS = _THEME["colors"]
_LAYOUT = _THEME["layout"]
NAVY = _COLORS["navy"]
BLUE = _COLORS["blue"]
GREEN = _COLORS["green"]
RED = _COLORS["red"]
AMBER = _COLORS["amber"]
LIGHT_BLUE = _COLORS["light_blue"]
LIGHT_GREEN = _COLORS["light_green"]
LIGHT_RED = _COLORS["light_red"]
LIGHT_AMBER = _COLORS["light_amber"]
LIGHT_GRAY = _COLORS["light_gray"]
BORDER = _COLORS["border"]
TEXT = _COLORS["text"]
MUTED = _COLORS["muted"]
WHITE = _COLORS["white"]


def _solid(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _thin_border() -> Border:
    side = Side(style="thin", color=BORDER)
    return Border(left=side, right=side, top=side, bottom=side)


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def _as_number(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def _normalise_header(value) -> str:
    return "".join(ch for ch in str(value or "").strip().lower() if ch.isalnum())


def _progress_columns(progress_ws) -> dict[str, int]:
    aliases = {
        "date": {"weekstart", "weeklydate", "date", "period", "week"},
        "plan": {"plan", "weeklyplan", "planned", "plannedprogress"},
        "actual": {"actual", "weeklyactual", "actualprogress"},
    }
    result: dict[str, int] = {}
    for col in range(1, progress_ws.max_column + 1):
        key = _normalise_header(progress_ws.cell(1, col).value)
        for role, names in aliases.items():
            if key in names and role not in result:
                result[role] = col
    missing = [role for role in ("date", "plan", "actual") if role not in result]
    if missing:
        headers = [str(progress_ws.cell(1, c).value or "") for c in range(1, progress_ws.max_column + 1)]
        raise RuntimeError(
            "Dashboard data source is missing required progress columns "
            f"{missing}. Found headers: {headers}"
        )
    return result


def _month_end(value: date) -> date:
    return date(value.year, value.month, monthrange(value.year, value.month)[1])


def _remove(workbook, name: str) -> None:
    if name in workbook.sheetnames:
        del workbook[name]


def _resolve_reference(workbook, value):
    if not isinstance(value, str) or not value.startswith("="):
        return value
    match = re.fullmatch(r"='?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)", value.strip())
    if not match:
        return value
    sheet_name, column, row = match.groups()
    if sheet_name not in workbook.sheetnames:
        return value
    return workbook[sheet_name][f"{column}{row}"].value


def _progress_rows(workbook, progress_ws) -> list[tuple[int, date]]:
    columns = _progress_columns(progress_ws)
    result: list[tuple[int, date]] = []
    for row in range(2, progress_ws.max_row + 1):
        raw_date = _resolve_reference(workbook, progress_ws.cell(row, columns["date"]).value)
        week_date = _as_date(raw_date)
        if week_date is None:
            continue
        result.append((row, week_date))
    if not result:
        raise RuntimeError(
            "Dashboard generation failed: progress sheet contains no usable weekly dates."
        )
    return result


def _monthly_groups(rows: list[tuple[int, date]]) -> list[tuple[date, list[int]]]:
    groups: dict[tuple[int, int], list[int]] = {}
    for row, value in rows:
        groups.setdefault((value.year, value.month), []).append(row)
    return [
        (_month_end(date(year, month, 1)), source_rows)
        for (year, month), source_rows in sorted(groups.items())
    ]


def _normalise_percent(value: float | None) -> float | None:
    if value is None:
        return None
    return value / 100.0 if abs(value) > 1 else value


def _source_percent_formula(progress_ws, row: int, col: int) -> str | float | None:
    value = progress_ws.cell(row, col).value
    numeric = _as_number(value)
    if numeric is not None:
        return _normalise_percent(numeric)
    column_letter = progress_ws.cell(1, col).column_letter
    return f'''=IF('{PROGRESS_SHEET}'!{column_letter}{row}="","",'{PROGRESS_SHEET}'!{column_letter}{row}/100)'''


def _build_data_sheet(workbook, progress_ws) -> None:
    _remove(workbook, DATA_SHEET)
    ws = workbook.create_sheet(DATA_SHEET)
    columns = _progress_columns(progress_ws)
    rows = _progress_rows(workbook, progress_ws)
    months = _monthly_groups(rows)

    ws.append(["Weekly Date", "Weekly Plan", "Weekly Actual", "Monthly Date", "Monthly Plan", "Monthly Actual", "Period", "Plan", "Actual"])
    for output_row, (source_row, week_date) in enumerate(rows, start=2):
        ws.cell(output_row, 1, week_date)
        ws.cell(output_row, 2, _source_percent_formula(progress_ws, source_row, columns["plan"]))
        ws.cell(output_row, 3, _source_percent_formula(progress_ws, source_row, columns["actual"]))
        ws.cell(output_row, 1).number_format = "dd/mm/yyyy"
        ws.cell(output_row, 2).number_format = "0.00%"
        ws.cell(output_row, 3).number_format = "0.00%"

    for output_row, (month_date, source_rows) in enumerate(months, start=2):
        last_row = source_rows[-1]
        ws.cell(output_row, 4, month_date)
        ws.cell(output_row, 5, _source_percent_formula(progress_ws, last_row, columns["plan"]))
        actual_letter = progress_ws.cell(1, columns["actual"]).column_letter
        actual_refs = ",".join(f"'{PROGRESS_SHEET}'!{actual_letter}{row}" for row in source_rows)
        ws.cell(output_row, 6, f'=IF(COUNT({actual_refs})=0,"",LOOKUP(2,1/(({actual_refs})<>""),({actual_refs}))/100)')
        ws.cell(output_row, 4).number_format = "mmm-yyyy"
        ws.cell(output_row, 5).number_format = "0.00%"
        ws.cell(output_row, 6).number_format = "0.00%"

    display_count = max(len(rows), len(months))
    for output_row in range(2, display_count + 2):
        weekly_row = output_row if output_row <= len(rows) + 1 else None
        monthly_row = output_row if output_row <= len(months) + 1 else None
        weekly_date = f"A{weekly_row}" if weekly_row else '""'
        weekly_plan = f"B{weekly_row}" if weekly_row else '""'
        weekly_actual = f"C{weekly_row}" if weekly_row else '""'
        monthly_date = f"D{monthly_row}" if monthly_row else '""'
        monthly_plan = f"E{monthly_row}" if monthly_row else '""'
        monthly_actual = f"F{monthly_row}" if monthly_row else '""'
        ws.cell(output_row, 7, f'=IF(Dashboard!$G$5="Weekly",{weekly_date},{monthly_date})')
        # Baseline plan is always rendered for the full project duration. Cutoff only
        # affects KPI calculation and the Actual curve.
        ws.cell(output_row, 8, f'=IF(G{output_row}="","",IF(Dashboard!$G$5="Weekly",{weekly_plan},{monthly_plan}))')
        ws.cell(output_row, 9, f'=IF(OR(G{output_row}="",G{output_row}>Dashboard!$K$5),"",IF(Dashboard!$G$5="Weekly",{weekly_actual},{monthly_actual}))')
        ws.cell(output_row, 7).number_format = "dd/mm/yyyy"
        ws.cell(output_row, 8).number_format = "0.00%"
        ws.cell(output_row, 9).number_format = "0.00%"

    if ws.max_row < 2 or ws["A2"].value is None:
        raise RuntimeError("Dashboard generation failed: Dashboard_Data contains no weekly rows.")
    ws.sheet_state = "hidden"


def _merge_title(ws, cell_range: str, value: str, size: int = 12) -> None:
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = value
    cell.font = Font(name=_FONT, bold=True, size=size, color=TEXT)
    cell.alignment = Alignment(vertical="center")


def _style_box(ws, cell_range: str, fill: str = WHITE) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.fill = _solid(fill)
            cell.border = _thin_border()


def _kpi(ws, title_range: str, value_range: str, title: str, formula: str, fill: str, color: str, number_format: str = "0.00%") -> None:
    _style_box(ws, f"{title_range.split(':')[0]}:{value_range.split(':')[1]}", fill)
    ws.merge_cells(title_range)
    ws.merge_cells(value_range)
    title_cell = ws[title_range.split(":")[0]]
    value_cell = ws[value_range.split(":")[0]]
    title_cell.value = title
    title_cell.font = Font(name=_FONT, size=9, color=MUTED)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell.value = formula
    value_cell.font = Font(name=_FONT, size=18, bold=True, color=color)
    value_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    value_cell.number_format = number_format


def _build_dashboard_sheet(workbook, project_name: str | None = None) -> None:
    _remove(workbook, DASHBOARD_SHEET)
    ws = workbook.create_sheet(DASHBOARD_SHEET, 0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.outlinePr.summaryBelow = True

    # Simple, wide dashboard canvas. Four KPI cards fit on one line.
    widths = {
        "A": 3, "B": 14, "C": 14, "D": 14,
        "E": 14, "F": 14, "G": 14,
        "H": 14, "I": 14, "J": 14,
        "K": 14, "L": 14, "M": 14,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row in range(1, 90):
        ws.row_dimensions[row].height = 20

    ws.merge_cells("B2:M2")
    ws["B2"] = _LAYOUT["title"]
    ws["B2"].font = Font(name=_FONT, size=20, bold=True, color=NAVY)
    ws["B2"].alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 32

    _merge_title(ws, "B4:M4", "PROJECT INFORMATION", 11)
    _style_box(ws, "B5:M6", LIGHT_GRAY)
    ws["B5"] = "Project"
    ws.merge_cells("C5:D5")
    ws["C5"] = project_name or "Project"
    ws["C5"].font = Font(name=_FONT, bold=True, color=TEXT)

    ws["F5"] = "View"
    ws.merge_cells("G5:H5")
    ws["G5"] = "Weekly"
    ws["J5"] = "Cutoff Date"
    ws.merge_cells("K5:M5")
    ws["K5"] = workbook[DATA_SHEET].cell(workbook[DATA_SHEET].max_row, 1).value
    ws["K5"].number_format = "dd/mm/yyyy"
    ws["G5"].font = Font(name=_FONT, bold=True, color=NAVY)
    ws["K5"].font = Font(name=_FONT, bold=True, color=NAVY)
    ws["G5"].alignment = Alignment(horizontal="center")
    ws["K5"].alignment = Alignment(horizontal="center")

    ws["B6"] = "Data source"
    ws.merge_cells("C6:H6")
    ws["C6"] = "Live formulas from progress / progress_table"
    ws["C6"].font = Font(name=_FONT, size=9, color=MUTED)
    ws["J6"] = "Chart rule"
    ws.merge_cells("K6:M6")
    ws["K6"] = "Plan: full baseline  |  Actual: to cutoff"
    ws["K6"].font = Font(name=_FONT, size=9, color=MUTED)

    view_validation = DataValidation(type="list", formula1='"Weekly,Monthly"', allow_blank=False)
    view_validation.error = "Choose Weekly or Monthly."
    view_validation.errorTitle = "Invalid reporting view"
    view_validation.prompt = "Switch the dashboard chart between Weekly and Monthly."
    view_validation.promptTitle = "Reporting View"
    view_validation.showInputMessage = True
    view_validation.showErrorMessage = True
    ws.add_data_validation(view_validation)
    view_validation.add(ws["G5"])

    last_week_row = max(2, len(_progress_rows(workbook, workbook[PROGRESS_SHEET])) + 1)
    cutoff_validation = DataValidation(
        type="list",
        formula1=f"'{DATA_SHEET}'!$A$2:$A${last_week_row}",
        allow_blank=False,
    )
    cutoff_validation.prompt = "Select the reporting cutoff date. KPI and Actual use this date."
    cutoff_validation.promptTitle = "Cutoff Date"
    cutoff_validation.showInputMessage = True
    ws.add_data_validation(cutoff_validation)
    cutoff_validation.add(ws["K5"])

    # KPI values always use the selected cutoff date regardless of chart view.
    _merge_title(ws, "B8:M8", "KPI SUMMARY", 11)
    plan_kpi = '=IFERROR(LOOKUP(2,1/((Dashboard_Data!$A$2:$A$250<=$K$5)*(Dashboard_Data!$B$2:$B$250<>"")),Dashboard_Data!$B$2:$B$250),0)'
    actual_kpi = '=IFERROR(LOOKUP(2,1/((Dashboard_Data!$A$2:$A$250<=$K$5)*(Dashboard_Data!$C$2:$C$250<>"")),Dashboard_Data!$C$2:$C$250),0)'
    _kpi(ws, "B9:D9", "B10:D12", "PLANNED PROGRESS", plan_kpi, LIGHT_BLUE, BLUE)
    _kpi(ws, "E9:G9", "E10:G12", "ACTUAL PROGRESS", actual_kpi, LIGHT_GREEN, GREEN)
    _kpi(ws, "H9:J9", "H10:J12", "SCHEDULE STATUS", '=IF(E10>=B10,"ON TRACK","DELAY "&TEXT(B10-E10,"0.00%"))', LIGHT_AMBER, AMBER, "General")
    _kpi(ws, "K9:M9", "K10:M12", "TIME IMPACT", '=IF(E10>=B10,"0 Days",MAX(0,$K$5-IFERROR(LOOKUP(E10,Dashboard_Data!$B$2:$B$250,Dashboard_Data!$A$2:$A$250),Dashboard_Data!$A$2))&" Days")', LIGHT_RED, RED, "General")
    for row in range(9, 13):
        ws.row_dimensions[row].height = 24
    ws.row_dimensions[10].height = 28
    ws.row_dimensions[11].height = 28
    ws.row_dimensions[12].height = 28

    # S-Curve: Plan is full baseline; Actual stops at selected cutoff.
    _merge_title(ws, "B15:M15", "S-CURVE — PLAN VS ACTUAL", 11)
    _style_box(ws, "B16:M34", WHITE)
    chart = LineChart()
    chart.title = None
    chart.style = int(_LAYOUT.get("chart_style", 2))
    chart.height = float(_LAYOUT["chart_height"])
    chart.width = float(_LAYOUT["chart_width"])
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.y_axis.majorUnit = float(_LAYOUT.get("chart_major_unit", 0.2))
    chart.y_axis.numFmt = "0%"
    chart.y_axis.title = "Progress"
    chart.x_axis.title = "Period"
    chart.legend.position = "t"

    # Make the plotting area lighter and cleaner than the default Excel chart.
    grid = GraphicalProperties()
    grid.line.solidFill = _COLORS.get("chart_grid", "E5E7EB")
    grid.line.width = 9000
    chart.y_axis.majorGridlines.graphicalProperties = grid
    axis_line = GraphicalProperties()
    axis_line.line.solidFill = _COLORS.get("chart_axis", "B8C2CC")
    axis_line.line.width = 9000
    chart.y_axis.spPr = axis_line
    chart.x_axis.spPr = axis_line

    max_rows = min(250, workbook[DATA_SHEET].max_row)
    data = Reference(workbook[DATA_SHEET], min_col=8, max_col=9, min_row=1, max_row=max_rows)
    cats = Reference(workbook[DATA_SHEET], min_col=7, min_row=2, max_row=max_rows)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    if len(chart.series) >= 2:
        chart.series[0].graphicalProperties.line.solidFill = BLUE
        chart.series[0].graphicalProperties.line.width = int(_LAYOUT.get("plan_line_width", 26000))
        chart.series[1].graphicalProperties.line.solidFill = GREEN
        chart.series[1].graphicalProperties.line.width = int(_LAYOUT.get("actual_line_width", 26000))
        if bool(_LAYOUT.get("chart_markers", False)):
            chart.series[0].marker.symbol = "circle"
            chart.series[0].marker.size = int(_LAYOUT.get("marker_size", 4))
            chart.series[1].marker.symbol = "circle"
            chart.series[1].marker.size = int(_LAYOUT.get("marker_size", 4))
    ws.add_chart(chart, "B16")

    ws.merge_cells("B35:M35")
    ws["B35"] = "Plan curve = full baseline duration    •    Actual curve = selected cutoff date"
    ws["B35"].font = Font(name=_FONT, size=9, italic=True, color=MUTED)
    ws["B35"].alignment = Alignment(horizontal="left", vertical="center")

    _merge_title(ws, "B37:M37", "ACTIVITY PROGRESS", 11)
    headers = ["WBS", "Activity", "Type", "Total", "Amount", "Progress"]
    starts = ["B", "C", "F", "G", "I", "K"]
    ends = ["B", "E", "F", "H", "J", "M"]
    for start, end, header in zip(starts, ends, headers):
        ws.merge_cells(f"{start}38:{end}38")
        cell = ws[f"{start}38"]
        cell.value = header
        cell.fill = _solid(NAVY)
        cell.font = Font(name=_FONT, bold=True, color=WHITE, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws[f"{start}38:{end}38"]:
            for c in row:
                c.border = _thin_border()

    source = workbook[TABLE_SHEET]
    output_row = 39
    source_row = 2
    shown = 0
    while source_row <= source.max_row and shown < int(_LAYOUT["activity_rows"]):
        plan_row = source_row
        actual_row = source_row + 1 if source_row + 1 <= source.max_row else source_row
        ws[f"B{output_row}"] = f"='{TABLE_SHEET}'!A{plan_row}"
        ws.merge_cells(f"C{output_row}:E{output_row}")
        ws[f"C{output_row}"] = f"='{TABLE_SHEET}'!B{plan_row}"
        ws[f"F{output_row}"] = "Plan"
        ws.merge_cells(f"G{output_row}:H{output_row}")
        ws[f"G{output_row}"] = f"='{TABLE_SHEET}'!C{plan_row}"
        ws.merge_cells(f"I{output_row}:J{output_row}")
        ws[f"I{output_row}"] = f'=IFERROR(G{output_row}*K{output_row},0)'
        ws.merge_cells(f"K{output_row}:M{output_row}")
        ws[f"K{output_row}"] = f'=IFERROR(SUMPRODUCT((\'{TABLE_SHEET}\'!$F$1:$ZZ$1<=$K$5)*\'{TABLE_SHEET}\'!$F{plan_row}:$ZZ{plan_row})/100,0)'

        actual_out = output_row + 1
        ws.merge_cells(f"C{actual_out}:E{actual_out}")
        ws[f"F{actual_out}"] = "Actual"
        ws.merge_cells(f"G{actual_out}:H{actual_out}")
        ws.merge_cells(f"I{actual_out}:J{actual_out}")
        ws[f"I{actual_out}"] = f'=IFERROR(G{output_row}*K{actual_out},0)'
        ws.merge_cells(f"K{actual_out}:M{actual_out}")
        ws[f"K{actual_out}"] = f'=IFERROR(SUMPRODUCT((\'{TABLE_SHEET}\'!$F$1:$ZZ$1<=$K$5)*\'{TABLE_SHEET}\'!$F{actual_row}:$ZZ{actual_row})/100,0)'

        for row_index in (output_row, actual_out):
            for row_cells in ws[f"B{row_index}:M{row_index}"]:
                for cell in row_cells:
                    cell.border = _thin_border()
                    cell.fill = _solid(WHITE if shown % 2 == 0 else LIGHT_GRAY)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws[f"G{row_index}"].number_format = "#,##0.00"
            ws[f"I{row_index}"].number_format = "#,##0.00"
            ws[f"K{row_index}"].number_format = "0.00%"
        ws[f"F{output_row}"].font = Font(name=_FONT, bold=True, color=BLUE)
        ws[f"F{actual_out}"].font = Font(name=_FONT, bold=True, color=GREEN)
        output_row += 2
        source_row += 2
        shown += 1

    if shown:
        ws.conditional_formatting.add(
            f"K39:K{output_row - 1}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=BLUE, showValue=True),
        )
    ws.auto_filter.ref = f"B38:M{max(38, output_row - 1)}"
    ws.print_area = f"B2:M{max(56, output_row - 1)}"


def build_dashboard(workbook, *, project_name: str | None = None) -> None:
    """Create the simple vertical dashboard as a separate first worksheet."""
    if PROGRESS_SHEET not in workbook.sheetnames or TABLE_SHEET not in workbook.sheetnames:
        return
    _build_data_sheet(workbook, workbook[PROGRESS_SHEET])
    _build_dashboard_sheet(workbook, project_name=project_name)


def build_dashboard_file(input_file: Path, output_file: Path, *, project_name: str | None = None) -> None:
    workbook = load_workbook(input_file, data_only=False)
    try:
        build_dashboard(workbook, project_name=project_name)
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(output_file)
    finally:
        workbook.close()
