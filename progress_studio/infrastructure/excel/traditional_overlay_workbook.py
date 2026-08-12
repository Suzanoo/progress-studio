from __future__ import annotations

from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

from progress_studio.domain.main_dataset import MainDataset
from progress_studio.infrastructure.excel.dashboard_workbook import BLUE, GREEN, DATA_SHEET

WEEKLY_OVERLAY_NAME = "PS_CURVE_OVERLAY_WEEKLY"
MONTHLY_OVERLAY_NAME = "PS_CURVE_OVERLAY_MONTHLY"
OVERLAY_TOP_ROW = 5
OVERLAY_MARKER_SIZE = 7
OVERLAY_LABEL_FORMAT = "0.0%"


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


def _overlay_chart(*, data_ws, date_col: int, plan_col: int, actual_col: int, last_row: int) -> LineChart:
    chart = LineChart()
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

    # Both chart and plot areas must be transparent so the underlying
    # timescale bars remain visible through the traditional overlay.
    chart.graphical_properties = GraphicalProperties(
        noFill=True, ln=LineProperties(noFill=True)
    )
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
        series.marker.size = OVERLAY_MARKER_SIZE
        series.marker.graphicalProperties.solidFill = color
        series.marker.graphicalProperties.line.solidFill = color
        # LW-12.3.1 experiment: show compact values at every marker.  We keep
        # category/series names off so labels stay as short percentages only.
        series.dLbls = DataLabelList(
            showVal=True,
            showCatName=False,
            showSerName=False,
            showLegendKey=False,
            numFmt=OVERLAY_LABEL_FORMAT,
            dLblPos="t",
        )
    return chart


def _responsive_anchor(*, first_col: int, last_col: int, top_row: int, bottom_row: int) -> TwoCellAnchor:
    """Create an Excel 'Move and size with cells' anchor over the schedule grid.

    AnchorMarker uses zero-based row/column indexes. The right/bottom marker is
    placed one cell beyond the intended range so the chart spans the complete
    Project Start -> Project Finish timescale and schedule height.
    """
    start = AnchorMarker(col=max(0, first_col - 1), row=max(0, top_row - 1))
    end = AnchorMarker(col=max(first_col, last_col), row=max(top_row, bottom_row))
    return TwoCellAnchor(editAs="twoCell", _from=start, to=end)


def build_traditional_overlays(workbook, dataset: MainDataset) -> tuple[bool, bool]:
    """LW-12.3.1: responsive all-marker overlays following the timescale grid."""
    ensure_overlay_visible_actual_columns(workbook)
    data_ws = workbook[DATA_SHEET]
    if not dataset.periods:
        return False, False

    first_timescale_col = dataset.periods[0].column
    last_weekly_col = dataset.periods[-1].column
    schedule_last_row = max(
        (row.row_number for row in dataset.rows),
        default=max(dataset.header_row + 1, OVERLAY_TOP_ROW + 1),
    )

    # Remove prior overlay charts when rebuilding the same workbook.
    for sheet_name in ("main", "main_monthly"):
        if sheet_name in workbook.sheetnames:
            workbook[sheet_name]._charts = []

    weekly_count = max(1, len(dataset.periods))
    weekly_last = min(data_ws.max_row, weekly_count + 1)
    monthly_count = sum(
        1 for r in range(2, data_ws.max_row + 1)
        if data_ws.cell(r, 4).value not in (None, "")
    )
    monthly_last = max(2, monthly_count + 1)
    last_monthly_col = first_timescale_col + max(1, monthly_count) - 1

    weekly_added = False
    if "main" in workbook.sheetnames:
        chart = _overlay_chart(
            data_ws=data_ws,
            date_col=1,
            plan_col=2,
            actual_col=16,
            last_row=weekly_last,
        )
        chart.anchor = _responsive_anchor(
            first_col=first_timescale_col,
            last_col=last_weekly_col,
            top_row=OVERLAY_TOP_ROW,
            bottom_row=schedule_last_row,
        )
        workbook["main"].add_chart(chart)
        weekly_added = True

    monthly_added = False
    if "main_monthly" in workbook.sheetnames:
        chart = _overlay_chart(
            data_ws=data_ws,
            date_col=4,
            plan_col=5,
            actual_col=17,
            last_row=monthly_last,
        )
        chart.anchor = _responsive_anchor(
            first_col=first_timescale_col,
            last_col=last_monthly_col,
            top_row=OVERLAY_TOP_ROW,
            bottom_row=schedule_last_row,
        )
        workbook["main_monthly"].add_chart(chart)
        monthly_added = True

    return weekly_added, monthly_added
