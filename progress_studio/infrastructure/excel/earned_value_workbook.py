from __future__ import annotations

from datetime import date, datetime

from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from progress_studio.domain.earned_value import EarnedValuePoint, EarnedValueResult


EARNED_VALUE_SHEET = "Earned Value"

_NAVY = "17365D"
_BLUE = "2F75B5"
_GREEN = "70AD47"
_RED = "C00000"
_LIGHT_BLUE = "EAF2F8"
_LIGHT_GRAY = "F3F4F6"
_BORDER = "D9DEE7"
_TEXT = "1F2937"
_MUTED = "667085"
_WHITE = "FFFFFF"


def _solid(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _thin_border() -> Border:
    side = Side(style="thin", color=_BORDER)
    return Border(left=side, right=side, top=side, bottom=side)


def _as_date(value: datetime | date | None) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    return value


def _latest_actual_point(
    result: EarnedValueResult,
) -> EarnedValuePoint | None:
    """Return the latest project point that belongs to the EV reporting horizon."""
    cutoff = _as_date(result.cutoff_date)
    latest: EarnedValuePoint | None = None
    for point in result.project_points:
        reporting_date = _as_date(point.reporting_date)
        if cutoff is not None and reporting_date is not None and reporting_date > cutoff:
            continue
        if point.earned_value is None:
            continue
        latest = point
    return latest


def _remove_existing_sheet(workbook) -> None:
    if EARNED_VALUE_SHEET in workbook.sheetnames:
        del workbook[EARNED_VALUE_SHEET]


def render_earned_value_sheet(workbook, result: EarnedValueResult) -> None:
    """Render EV-2 project representation from an already-derived EV-1 result.

    This function owns presentation only. It does not calculate BAC/PV/EV and
    intentionally does not render BOQ detail; BOQ representation belongs to EV-3.
    """
    _remove_existing_sheet(workbook)
    ws = workbook.create_sheet(EARNED_VALUE_SHEET)

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A12"

    # Title
    ws.merge_cells("A1:H2")
    title = ws["A1"]
    title.value = "EARNED VALUE"
    title.font = Font(name="Aptos Display", size=20, bold=True, color=_WHITE)
    title.fill = _solid(_NAVY)
    title.alignment = Alignment(vertical="center")
    for row in ws["A1:H2"]:
        for cell in row:
            cell.fill = _solid(_NAVY)

    latest = _latest_actual_point(result)

    # Status / basis
    ws["A4"] = "Status Date"
    ws["B4"] = result.cutoff_date
    ws["B4"].number_format = "dd-mmm-yyyy"
    ws["D4"] = "Budget Basis"
    ws["E4"] = "Mapped BOQ Amount"

    # KPI cards
    kpis = (
        ("BAC", result.project_bac, "#,##0.00"),
        ("PV", None if latest is None else latest.planned_value, "#,##0.00"),
        ("EV", None if latest is None else latest.earned_value, "#,##0.00"),
        ("SV", None if latest is None else latest.schedule_variance, "#,##0.00"),
        ("SPI", None if latest is None else latest.schedule_performance_index, "0.00"),
    )
    start_cols = (1, 3, 5, 7, 9)
    for (label, value, number_format), col in zip(kpis, start_cols):
        ws.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col + 1)
        ws.merge_cells(start_row=7, start_column=col, end_row=8, end_column=col + 1)
        label_cell = ws.cell(6, col)
        value_cell = ws.cell(7, col)
        label_cell.value = label
        label_cell.font = Font(name="Aptos", size=10, bold=True, color=_MUTED)
        label_cell.alignment = Alignment(horizontal="center")
        label_cell.fill = _solid(_LIGHT_GRAY)
        value_cell.value = value
        value_cell.number_format = number_format
        value_cell.font = Font(name="Aptos Display", size=16, bold=True, color=_TEXT)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell.fill = _solid(_WHITE)
        for row in range(6, 9):
            for c in range(col, col + 2):
                ws.cell(row, c).border = _thin_border()

    # Period table
    headers = ("Period", "Reporting Date", "PV", "EV", "SV", "SPI")
    header_row = 11
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(header_row, col, header)
        cell.font = Font(name="Aptos", bold=True, color=_WHITE)
        cell.fill = _solid(_NAVY)
        cell.alignment = Alignment(horizontal="center")
        cell.border = _thin_border()

    first_data_row = header_row + 1
    for row_index, point in enumerate(result.project_points, start=first_data_row):
        values = (
            point.period_key,
            point.reporting_date,
            point.planned_value,
            point.earned_value,
            point.schedule_variance,
            point.schedule_performance_index,
        )
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row_index, col, value)
            cell.border = _thin_border()
            if row_index % 2 == 0:
                cell.fill = _solid(_LIGHT_BLUE)
        ws.cell(row_index, 2).number_format = "dd-mmm-yyyy"
        for col in (3, 4, 5):
            ws.cell(row_index, col).number_format = "#,##0.00"
        ws.cell(row_index, 6).number_format = "0.00"

    last_data_row = first_data_row + len(result.project_points) - 1

    # PV / EV S-curve. EV cells after cutoff are blank by EV-1 contract, so the
    # earned curve naturally stops while PV may continue to project finish.
    if result.project_points:
        chart = LineChart()
        chart.title = "Planned Value vs Earned Value"
        chart.style = 10
        chart.height = 9.0
        chart.width = 17.5
        chart.y_axis.title = "Value"
        chart.x_axis = DateAxis(crosses="autoZero")
        chart.x_axis.title = "Reporting Date"
        chart.x_axis.number_format = "mmm-yy"
        chart.legend.position = "b"

        data = Reference(
            ws,
            min_col=3,
            max_col=4,
            min_row=header_row,
            max_row=last_data_row,
        )
        dates = Reference(
            ws,
            min_col=2,
            min_row=first_data_row,
            max_row=last_data_row,
        )
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(dates)

        if len(chart.series) >= 1:
            chart.series[0].graphicalProperties.line.solidFill = _BLUE
            chart.series[0].graphicalProperties.line.width = 24000
        if len(chart.series) >= 2:
            chart.series[1].graphicalProperties.line.solidFill = _GREEN
            chart.series[1].graphicalProperties.line.width = 24000

        ws.add_chart(chart, "H11")

    # Sizing
    widths = {
        "A": 16,
        "B": 18,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 12,
        "G": 3,
        "H": 14,
        "I": 14,
        "J": 14,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    for row in range(1, max(last_data_row, 12) + 1):
        ws.row_dimensions[row].height = 20
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[7].height = 28
    ws.row_dimensions[8].height = 28

    # Keep summary labels restrained.
    for coord in ("A4", "D4"):
        ws[coord].font = Font(name="Aptos", size=10, bold=True, color=_MUTED)
    for coord in ("B4", "E4"):
        ws[coord].font = Font(name="Aptos", size=10, color=_TEXT)
