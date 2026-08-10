
from __future__ import annotations

from datetime import date, datetime

from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from progress_studio.domain.activity_table import ActivityTableModel
from progress_studio.domain.main_dataset import MainDataset
from progress_studio.domain.progress_cache import ProgressCache
from progress_studio.infrastructure.excel.dashboard_workbook import (
    AMBER,
    BLUE,
    BORDER,
    DATA_SHEET,
    DASHBOARD_SHEET,
    GREEN,
    LIGHT_AMBER,
    LIGHT_BLUE,
    LIGHT_GRAY,
    LIGHT_GREEN,
    LIGHT_RED,
    MUTED,
    NAVY,
    RED,
    TEXT,
    WHITE,
    _FONT,
    _LAYOUT,
    _merge_title,
    _remove,
    _solid,
    _style_box,
    _thin_border,
)
from progress_studio.services.activity_table_deriver import ActivityTableDeriver
from progress_studio.services.progress_cache_deriver import ProgressCacheDeriver


def _as_date(value: date | datetime | None) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    return value


def _default_cutoff(cache: ProgressCache) -> date | None:
    latest_actual = None
    latest_any = None
    for point in cache.points:
        point_date = _as_date(point.reporting_date)
        if point_date is not None:
            latest_any = point_date
            if point.actual_cumulative is not None:
                latest_actual = point_date
    return latest_actual or latest_any


def _project_dates(dataset: MainDataset) -> tuple[date | None, date | None]:
    starts = [row.plan_start.date() for row in dataset.activities if row.plan_start is not None]
    finishes = [row.plan_finish.date() for row in dataset.activities if row.plan_finish is not None]
    return (min(starts) if starts else None, max(finishes) if finishes else None)


def _value_at_cutoff(cache: ProgressCache, cutoff: date | None, attr: str) -> float:
    value = 0.0
    for point in cache.points:
        point_date = _as_date(point.reporting_date)
        if cutoff is not None and point_date is not None and point_date > cutoff:
            continue
        candidate = getattr(point, attr)
        if candidate is not None:
            value = float(candidate)
    return value


def _build_live_data_sheet(workbook, cache: ProgressCache, cutoff: date | None) -> None:
    _remove(workbook, DATA_SHEET)
    ws = workbook.create_sheet(DATA_SHEET)
    ws.append(["Date", "Plan", "Actual"])
    for point in cache.points:
        point_date = _as_date(point.reporting_date)
        actual = point.actual_cumulative
        if cutoff is not None and point_date is not None and point_date > cutoff:
            actual = None
        ws.append([point.reporting_date, point.plan_cumulative, actual])

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 1).number_format = "dd/mm/yyyy"
        ws.cell(row, 2).number_format = "0.00%"
        ws.cell(row, 3).number_format = "0.00%"
    ws.sheet_state = "hidden"


def _kpi_box(ws, title_range: str, value_range: str, title: str, value, fill: str, color: str, number_format: str = "0.00%") -> None:
    ws.merge_cells(title_range)
    ws[title_range.split(":")[0]] = title
    ws[title_range.split(":")[0]].font = Font(name=_FONT, bold=True, color=color, size=9)
    ws[title_range.split(":")[0]].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(value_range)
    cell = ws[value_range.split(":")[0]]
    cell.value = value
    cell.number_format = number_format
    cell.fill = _solid(fill)
    cell.font = Font(name=_FONT, bold=True, color=color, size=15)
    cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws[title_range]:
        for c in row:
            c.fill = _solid(fill)
            c.border = _thin_border()
    for row in ws[value_range]:
        for c in row:
            c.fill = _solid(fill)
            c.border = _thin_border()


def _write_activity_section(ws, model: ActivityTableModel) -> None:
    _merge_title(ws, "B37:M37", "ACTIVITY PROGRESS", 11)
    ws.merge_cells("N37:O37")
    ws["N37"] = "Status"
    ws["N37"].font = Font(name=_FONT, bold=True, color=NAVY, size=9)
    ws["N37"].alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("P37:Q37")
    ws["P37"] = "All"
    ws["P37"].fill = _solid(LIGHT_BLUE)
    ws["P37"].font = Font(name=_FONT, bold=True, color=NAVY, size=9)
    ws["P37"].alignment = Alignment(horizontal="center", vertical="center")
    validation = DataValidation(
        type="list",
        formula1='"All,Behind,On Track,Complete,Not Started"',
        allow_blank=False,
    )
    ws.add_data_validation(validation)
    validation.add(ws["P37"])

    headers = ["WBS", "Activity", "Type", "Total", "Amount", "Progress", "Variance", "Status"]
    starts = ["B", "C", "F", "H", "J", "L", "N", "P"]
    ends = ["B", "E", "G", "I", "K", "M", "O", "Q"]
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

    output_row = 39
    pair_index = 0
    for item in model.rows:
        row = output_row
        is_plan = item.type_label == "Plan"
        if is_plan:
            ws[f"B{row}"] = item.wbs if item.row_type != "project summary" else "PROJECT"
            ws.merge_cells(f"C{row}:E{row}")
            ws[f"C{row}"] = item.activity
        else:
            ws.merge_cells(f"C{row}:E{row}")

        ws.merge_cells(f"F{row}:G{row}")
        ws[f"F{row}"] = item.type_label
        ws.merge_cells(f"H{row}:I{row}")
        ws[f"H{row}"] = item.total
        ws.merge_cells(f"J{row}:K{row}")
        ws[f"J{row}"] = item.amount
        ws.merge_cells(f"L{row}:M{row}")
        ws[f"L{row}"] = item.progress
        ws.merge_cells(f"N{row}:O{row}")
        ws[f"N{row}"] = item.variance
        ws.merge_cells(f"P{row}:Q{row}")
        ws[f"P{row}"] = item.status

        base_fill = WHITE if pair_index % 2 == 0 else LIGHT_GRAY
        if item.row_type in {"project summary", "wbs"}:
            base_fill = LIGHT_BLUE if item.outline_level <= 1 else "F5F8FC"

        for row_cells in ws[f"B{row}:Q{row}"]:
            for cell in row_cells:
                cell.border = _thin_border()
                cell.fill = _solid(base_fill)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if item.row_type in {"project summary", "wbs"}:
                    cell.font = Font(name=_FONT, bold=True, color=NAVY, size=9)

        if is_plan:
            ws[f"C{row}"].alignment = Alignment(
                vertical="center",
                wrap_text=True,
                indent=min(max(item.outline_level, 0), 4),
            )

        pa_fill = LIGHT_BLUE if is_plan else LIGHT_GREEN
        pa_color = BLUE if is_plan else GREEN
        ws[f"F{row}"].fill = _solid(pa_fill)
        ws[f"F{row}"].font = Font(name=_FONT, bold=True, color=pa_color, size=9)
        ws[f"L{row}"].fill = _solid(pa_fill)
        ws[f"L{row}"].font = Font(name=_FONT, bold=True, color=pa_color, size=9)

        ws.row_dimensions[row].outlineLevel = min(max(item.outline_level, 0), 7)
        ws.row_dimensions[row].height = 22
        ws[f"H{row}"].number_format = "#,##0.00"
        ws[f"J{row}"].number_format = "#,##0.00"
        ws[f"L{row}"].number_format = "0.00%"
        ws[f"N{row}"].number_format = "0.00%;[Red]-0.00%;0.00%"

        output_row += 1
        if not is_plan:
            pair_index += 1

    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.print_area = f"B2:Q{max(56, output_row - 1)}"


def build_live_dashboard(
    workbook,
    dataset: MainDataset,
    *,
    project_name: str | None = None,
    cutoff: date | datetime | None = None,
) -> None:
    """Render the LW-5 Dashboard without `progress` or `progress_table` sheets."""
    cache = ProgressCacheDeriver().derive(dataset)
    cutoff_date = _as_date(cutoff) or _default_cutoff(cache)
    activity_model = ActivityTableDeriver().derive(dataset, cutoff=cutoff_date)
    _build_live_data_sheet(workbook, cache, cutoff_date)

    _remove(workbook, DASHBOARD_SHEET)
    ws = workbook.create_sheet(DASHBOARD_SHEET, 0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"

    widths = {
        "A": 3, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14,
        "H": 14, "I": 14, "J": 14, "K": 14, "L": 14, "M": 14,
        "N": 11, "O": 11, "P": 12, "Q": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.merge_cells("B2:M2")
    ws["B2"] = _LAYOUT["title"]
    ws["B2"].font = Font(name=_FONT, size=20, bold=True, color=NAVY)

    _merge_title(ws, "B4:M4", "PROJECT INFORMATION", 11)
    _style_box(ws, "B5:M6", LIGHT_GRAY)
    ws["B5"] = "Project"
    ws.merge_cells("C5:D5")
    ws["C5"] = project_name or dataset.workbook_name
    ws["F5"] = "View"
    ws.merge_cells("G5:H5")
    ws["G5"] = "Weekly"
    ws["J5"] = "Cutoff Date"
    ws.merge_cells("K5:M5")
    ws["K5"] = cutoff_date
    ws["K5"].number_format = "dd/mm/yyyy"

    ws["B6"] = "Data source"
    ws.merge_cells("C6:H6")
    ws["C6"] = "Live: MainDataset → Progress Cache / Activity Deriver"
    ws["C6"].font = Font(name=_FONT, size=9, color=MUTED)
    ws["J6"] = "LW-5"
    ws.merge_cells("K6:M6")
    ws["K6"] = "Weekly dashboard contract • Monthly follows in LW-6"
    ws["K6"].font = Font(name=_FONT, size=9, color=MUTED)

    plan_value = _value_at_cutoff(cache, cutoff_date, "plan_cumulative")
    actual_value = _value_at_cutoff(cache, cutoff_date, "actual_cumulative")
    variance = actual_value - plan_value
    status = "ON SCHEDULE" if abs(variance) < 1e-12 else ("DELAY" if variance < 0 else "AHEAD")
    start_date, finish_date = _project_dates(dataset)
    duration_days = max(0, (finish_date - start_date).days) if start_date and finish_date else 0
    impact_days = round(abs(variance) * duration_days)

    _merge_title(ws, "B8:M8", "KPI SUMMARY", 11)
    _kpi_box(ws, "B9:D9", "B10:D12", "PLANNED PROGRESS", plan_value, LIGHT_BLUE, BLUE)
    _kpi_box(ws, "E9:G9", "E10:G12", "ACTUAL PROGRESS", actual_value, LIGHT_GREEN, GREEN)
    _kpi_box(ws, "H9:J9", "H10:J12", "SCHEDULE STATUS", status, LIGHT_AMBER if variance >= 0 else LIGHT_RED, AMBER if variance >= 0 else RED, "General")
    _kpi_box(ws, "K9:M9", "K10:M12", "TIME IMPACT", f"{impact_days} Days", LIGHT_GREEN if variance >= 0 else LIGHT_RED, GREEN if variance >= 0 else RED, "General")

    _merge_title(ws, "B15:M15", "S-CURVE — PLAN VS ACTUAL", 11)
    _style_box(ws, "B16:M34", WHITE)
    data_ws = workbook[DATA_SHEET]
    chart = LineChart()
    chart.height = float(_LAYOUT["chart_height"])
    chart.width = float(_LAYOUT["chart_width"])
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.y_axis.numFmt = "0%"
    chart.legend.position = "t"
    data = Reference(data_ws, min_col=2, max_col=3, min_row=1, max_row=data_ws.max_row)
    cats = Reference(data_ws, min_col=1, min_row=2, max_row=data_ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    if len(chart.series) >= 2:
        chart.series[0].graphicalProperties.line.solidFill = BLUE
        chart.series[1].graphicalProperties.line.solidFill = GREEN
    ws.add_chart(chart, "B16")

    ws.merge_cells("B35:M35")
    ws["B35"] = "Plan curve = full baseline duration    •    Actual curve = selected cutoff snapshot"
    ws["B35"].font = Font(name=_FONT, size=9, italic=True, color=MUTED)

    _write_activity_section(ws, activity_model)
