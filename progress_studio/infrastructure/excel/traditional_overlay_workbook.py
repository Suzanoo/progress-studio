from __future__ import annotations

from openpyxl.chart import LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.utils import get_column_letter

from progress_studio.domain.main_dataset import MainDataset
from progress_studio.infrastructure.excel.dashboard_workbook import BLUE, GREEN, DATA_SHEET

WEEKLY_OVERLAY_NAME = "PS_CURVE_OVERLAY_WEEKLY"
MONTHLY_OVERLAY_NAME = "PS_CURVE_OVERLAY_MONTHLY"
OVERLAY_ANCHOR_ROW = 5
OVERLAY_HEIGHT_CM = 12.0
OVERLAY_MIN_WIDTH_CM = 18.0
OVERLAY_MAX_WIDTH_CM = 42.0
OVERLAY_PERIOD_WIDTH_CM = 0.55


def ensure_overlay_visible_actual_columns(workbook) -> None:
    """Reuse the Dashboard cutoff rule for fixed Weekly/Monthly overlay views."""
    if DATA_SHEET not in workbook.sheetnames:
        raise ValueError("Dashboard_Data is required before building traditional overlays.")
    ws = workbook[DATA_SHEET]
    ws["P1"] = "Weekly Actual Visible"
    ws["Q1"] = "Monthly Actual Visible"
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 16, f'=IF(A{row}="",NA(),IF(A{row}>Dashboard!$K$5,NA(),IF(C{row}="",NA(),C{row})))')
        ws.cell(row, 17, f'=IF(D{row}="",NA(),IF(D{row}>Dashboard!$K$5,NA(),IF(F{row}="",NA(),F{row})))')
        ws.cell(row, 16).number_format = "0.00%"
        ws.cell(row, 17).number_format = "0.00%"


def _overlay_chart(*, data_ws, date_col: int, plan_col: int, actual_col: int, last_row: int, period_count: int) -> LineChart:
    chart = LineChart()
    chart.height = OVERLAY_HEIGHT_CM
    # Prototype geometry: roughly follows visible timescale width, but is capped
    # so very long schedules remain manageable in Excel.
    chart.width = max(OVERLAY_MIN_WIDTH_CM, min(OVERLAY_MAX_WIDTH_CM, period_count * OVERLAY_PERIOD_WIDTH_CM))
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.y_axis.majorUnit = 0.25
    chart.y_axis.numFmt = "0%"
    chart.y_axis.title = None
    chart.x_axis.title = None
    chart.x_axis.tickLblPos = "none"
    chart.legend.position = "t"
    chart.display_blanks = "gap"
    chart.y_axis.majorGridlines = None

    # LW-12.3/12.4: both layers must be transparent.  Excel charts have an
    # outer chart area and a separate inner plot area; making only the outer
    # layer transparent still hides schedule bars behind the plot rectangle.
    transparent = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    chart.graphical_properties = transparent
    chart.plot_area.graphicalProperties = GraphicalProperties(
        noFill=True, ln=LineProperties(noFill=True)
    )

    plan = Reference(data_ws, min_col=plan_col, max_col=plan_col, min_row=1, max_row=last_row)
    actual = Reference(data_ws, min_col=actual_col, max_col=actual_col, min_row=1, max_row=last_row)
    cats = Reference(data_ws, min_col=date_col, min_row=2, max_row=last_row)
    chart.add_data(plan, titles_from_data=True)
    chart.add_data(actual, titles_from_data=True)
    chart.set_categories(cats)

    for series, color in zip(chart.series[:2], (BLUE, GREEN)):
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 19050
        series.marker.symbol = "circle"
        series.marker.size = 5
        series.marker.graphicalProperties.solidFill = color
        series.marker.graphicalProperties.line.solidFill = color
    return chart


def build_traditional_overlays(workbook, dataset: MainDataset) -> tuple[bool, bool]:
    """LW-12.1-12.4: fixed all-marker overlays safe for outline grouping."""
    ensure_overlay_visible_actual_columns(workbook)
    data_ws = workbook[DATA_SHEET]
    first_timescale_col = dataset.periods[0].column if dataset.periods else None
    if first_timescale_col is None:
        return False, False

    # Remove prior overlay charts when rebuilding the same workbook.
    for sheet_name in ("main", "main_monthly"):
        if sheet_name in workbook.sheetnames:
            workbook[sheet_name]._charts = []

    weekly_count = max(1, len(dataset.periods))
    weekly_last = min(data_ws.max_row, weekly_count + 1)
    monthly_count = sum(1 for r in range(2, data_ws.max_row + 1) if data_ws.cell(r, 4).value not in (None, ""))
    monthly_last = max(2, monthly_count + 1)
    # Anchor in the ungrouped header zone. openpyxl writes a oneCellAnchor,
    # so the chart keeps a fixed extent; outline collapse/expand below row 5
    # does not resize the overlay or move its timescale origin.
    anchor = f"{get_column_letter(first_timescale_col)}{OVERLAY_ANCHOR_ROW}"

    weekly_added = False
    if "main" in workbook.sheetnames:
        chart = _overlay_chart(
            data_ws=data_ws, date_col=1, plan_col=2, actual_col=16,
            last_row=weekly_last, period_count=weekly_count,
        )
        workbook["main"].add_chart(chart, anchor)
        weekly_added = True

    monthly_added = False
    if "main_monthly" in workbook.sheetnames:
        chart = _overlay_chart(
            data_ws=data_ws, date_col=4, plan_col=5, actual_col=17,
            last_row=monthly_last, period_count=max(1, monthly_count),
        )
        workbook["main_monthly"].add_chart(chart, anchor)
        monthly_added = True

    return weekly_added, monthly_added
