from __future__ import annotations

from datetime import date, datetime

from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.text import RichText
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

from progress_studio.domain.main_dataset import MainDataset
from progress_studio.infrastructure.excel.dashboard_workbook import BLUE, GREEN, DATA_SHEET

WEEKLY_OVERLAY_NAME = "PS_CURVE_OVERLAY_WEEKLY"
MONTHLY_OVERLAY_NAME = "PS_CURVE_OVERLAY_MONTHLY"
OVERLAY_TOP_ROW = 5
CUTOFF_LABEL_FILL = "1F4E78"
CUTOFF_VALUE_FILL = "D9EAF7"
OVERLAY_MARKER_SIZE = 7
OVERLAY_LABEL_FORMAT = "0.0%"
OVERLAY_LABEL_FONT_SIZE = 700  # DrawingML uses 1/100 pt -> 7 pt.
PLAN_LABEL_TEXT = "1F4E79"
PLAN_LABEL_FILL = "DDEBF7"
PLAN_LABEL_BORDER = "9CC2E5"
ACTUAL_LABEL_TEXT = "385723"
ACTUAL_LABEL_FILL = "E2F0D9"
ACTUAL_LABEL_BORDER = "A9D18E"
CUTOFF_RED = "C00000"
CUTOFF_LABEL_BG = "FCE4D6"
CUTOFF_LABEL_BORDER = "C00000"


def ensure_overlay_visible_actual_columns(
    workbook,
    *,
    weekly_cutoff_ref: str,
    monthly_cutoff_ref: str,
) -> None:
    """Build independent cutoff-aware Actual helpers and cutoff-line helpers.

    Dashboard, main, and main_monthly deliberately own separate cutoff state.
    The traditional overlays never read Dashboard!K5 after their initial values
    are seeded; each view masks Actual and renders its red cutoff line from its
    own local cutoff cell.
    """
    if DATA_SHEET not in workbook.sheetnames:
        raise ValueError("Dashboard_Data is required before building traditional overlays.")
    ws = workbook[DATA_SHEET]
    ws["P1"] = "Weekly Actual Visible"
    ws["Q1"] = "Monthly Actual Visible"
    ws["R1"] = "Weekly Cutoff Line"
    ws["S1"] = "Monthly Cutoff Line"
    for row in range(2, ws.max_row + 1):
        next_row = row + 1
        ws.cell(row, 16, f'=IF(A{row}="",NA(),IF(A{row}>{weekly_cutoff_ref},NA(),IF(C{row}="",NA(),C{row})))')
        ws.cell(row, 17, f'=IF(D{row}="",NA(),IF(D{row}>{monthly_cutoff_ref},NA(),IF(F{row}="",NA(),F{row})))')
        # A single 1.0 point plus a full-height negative Y error bar produces
        # a vertical cutoff line without VBA or movable worksheet shapes.
        ws.cell(row, 18, f'=IF(A{row}="",NA(),IF(AND(A{row}<={weekly_cutoff_ref},OR(A{next_row}="",A{next_row}>{weekly_cutoff_ref})),1,NA()))')
        ws.cell(
            row, 19,
            f'=IF(D{row}="",NA(),IF(AND(D{row}<={monthly_cutoff_ref},OR(D{next_row}="",D{next_row}>{monthly_cutoff_ref})),1,NA()))'
        )
        ws.cell(row, 16).number_format = "0.00%"
        ws.cell(row, 17).number_format = "0.00%"
        ws.cell(row, 18).number_format = "0%"
        ws.cell(row, 19).number_format = "0%"



def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None

def _project_dates(dataset: MainDataset):
    rows = dataset.activities or tuple(
        row for row in dataset.rows if row.plan_start is not None or row.plan_finish is not None
    )
    starts = [row.plan_start for row in rows if row.plan_start is not None]
    finishes = [row.plan_finish for row in rows if row.plan_finish is not None]
    return (min(starts) if starts else None, max(finishes) if finishes else None)


def _weekly_project_window(dataset: MainDataset) -> tuple[int, int, int, int]:
    """Return Dashboard_Data row bounds and physical timescale columns for project duration."""
    start, finish = _project_dates(dataset)
    dated = [(idx, p) for idx, p in enumerate(dataset.periods) if p.reporting_date is not None]
    if not dated or start is None or finish is None:
        return 2, max(2, len(dataset.periods) + 1), dataset.periods[0].column, dataset.periods[-1].column

    first_idx = next((idx for idx, p in dated if p.reporting_date >= start), dated[0][0])
    last_idx = next((idx for idx, p in dated if p.reporting_date >= finish), dated[-1][0])
    if last_idx < first_idx:
        last_idx = first_idx
    return (
        first_idx + 2,
        last_idx + 2,
        dataset.periods[first_idx].column,
        dataset.periods[last_idx].column,
    )


def _monthly_project_window(data_ws, dataset: MainDataset) -> tuple[int, int, int, int]:
    """Return monthly helper row bounds and physical monthly columns for project duration."""
    start, finish = _project_dates(dataset)
    months = []
    for row in range(2, data_ws.max_row + 1):
        value = data_ws.cell(row, 4).value
        if value in (None, ""):
            continue
        months.append((row, value))
    first_timescale_col = dataset.periods[0].column
    if not months or start is None or finish is None:
        last = months[-1][0] if months else 2
        return 2, last, first_timescale_col, first_timescale_col + max(0, len(months) - 1)

    start_key = (start.year, start.month)
    finish_key = (finish.year, finish.month)
    first_pos = next((i for i, (_, d) in enumerate(months) if (d.year, d.month) >= start_key), 0)
    last_pos = max(
        (i for i, (_, d) in enumerate(months) if (d.year, d.month) <= finish_key),
        default=len(months) - 1,
    )
    if last_pos < first_pos:
        last_pos = first_pos
    return (
        months[first_pos][0],
        months[last_pos][0],
        first_timescale_col + first_pos,
        first_timescale_col + last_pos,
    )


def _scurve_plan_row(dataset: MainDataset) -> int:
    candidates = [
        row.row_number for row in dataset.rows
        if row.row_type.strip().lower() == "s-curve"
        and row.description.strip().lower() == "plan"
    ]
    if candidates:
        return min(candidates)
    any_scurve = [row.row_number for row in dataset.rows if row.row_type.strip().lower() == "s-curve"]
    if any_scurve:
        return min(any_scurve)
    return max((row.row_number for row in dataset.rows), default=dataset.header_row + 2) + 1


def _add_cutoff_control(
    ws,
    dataset: MainDataset,
    *,
    row: int,
    list_col: str,
    list_last_row: int,
    initial_value,
    display_format: str,
) -> str:
    """Add a compact independent cutoff selector in columns L:M.

    Column M is intentionally the editable value cell on both traditional
    views.  Keeping the control out of the timescale footprint makes it easy
    to reach even when the transparent overlay is visible.
    """
    label_col = 12  # L
    value_col = 13  # M
    label = ws.cell(row, label_col)
    value = ws.cell(row, value_col)
    label.value = "Cutoff Date"
    label.fill = PatternFill("solid", fgColor=CUTOFF_LABEL_FILL)
    label.font = Font(color="FFFFFF", bold=True)
    label.alignment = Alignment(horizontal="right", vertical="center")
    value.value = initial_value
    value.number_format = display_format
    value.fill = PatternFill("solid", fgColor=CUTOFF_VALUE_FILL)
    value.font = Font(color="1F1F1F", bold=True)
    value.alignment = Alignment(horizontal="center", vertical="center")
    value.comment = Comment(
        "This cutoff belongs to this sheet only. Press F9 or Save after changing it to recalculate the overlay.",
        "Progress Studio",
    )
    validation = DataValidation(
        type="list",
        formula1=f'=INDIRECT("{DATA_SHEET}!${list_col}$2:${list_col}${max(2, list_last_row)}")',
        allow_blank=False,
    )
    validation.promptTitle = "Cutoff Date"
    validation.prompt = "Choose this sheet's reporting cutoff date."
    validation.errorTitle = "Invalid Cutoff"
    validation.error = "Select a date from the list."
    validation.errorStyle = "stop"
    validation.showInputMessage = True
    validation.showErrorMessage = True
    ws.add_data_validation(validation)
    validation.add(value)
    return f"'{ws.title}'!${get_column_letter(value_col)}${row}"


def _label_text_properties(text_color: str) -> RichText:
    """Compact 7 pt series-tinted label text for a dense traditional overlay."""
    run = CharacterProperties(sz=OVERLAY_LABEL_FONT_SIZE, solidFill=text_color)
    paragraph = Paragraph(pPr=ParagraphProperties(defRPr=run))
    return RichText(bodyPr=RichTextProperties(), p=[paragraph])


def _label_graphical_properties(fill_color: str, border_color: str) -> GraphicalProperties:
    """Pale opaque tag background so values remain readable over schedule bars."""
    props = GraphicalProperties(solidFill=fill_color)
    props.line.solidFill = border_color
    props.line.width = 6350  # ~0.5 pt; enough separation without a heavy box.
    return props


def _overlay_chart(*, data_ws, date_col: int, plan_col: int, actual_col: int, cutoff_col: int | None = None, first_row: int, last_row: int) -> LineChart:
    chart = LineChart()
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.y_axis.majorUnit = 0.25
    chart.y_axis.numFmt = "0%"
    chart.y_axis.title = None
    chart.x_axis.title = None
    chart.x_axis.tickLblPos = "none"
    chart.legend = None
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

    plan = Reference(data_ws, min_col=plan_col, max_col=plan_col, min_row=first_row, max_row=last_row)
    actual = Reference(data_ws, min_col=actual_col, max_col=actual_col, min_row=first_row, max_row=last_row)
    cats = Reference(data_ws, min_col=date_col, min_row=first_row, max_row=last_row)
    chart.add_data(plan, titles_from_data=False)
    chart.add_data(actual, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].tx = SeriesLabel(v="Plan")
    chart.series[1].tx = SeriesLabel(v="Actual")

    label_styles = (
        (BLUE, "t", PLAN_LABEL_TEXT, PLAN_LABEL_FILL, PLAN_LABEL_BORDER),
        (GREEN, "b", ACTUAL_LABEL_TEXT, ACTUAL_LABEL_FILL, ACTUAL_LABEL_BORDER),
    )
    for series, (color, position, text_color, fill_color, border_color) in zip(
        chart.series[:2], label_styles
    ):
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 19050
        series.marker.symbol = "circle"
        series.marker.size = OVERLAY_MARKER_SIZE
        series.marker.graphicalProperties.solidFill = color
        series.marker.graphicalProperties.line.solidFill = color
        # LW-12.3.2: keep every value, but turn labels into compact tinted tags.
        # Plan sits above its curve; Actual sits below so nearby series do not
        # compete for the same vertical space.  Pale backgrounds preserve
        # readability while keeping the schedule bars visible around the tags.
        series.dLbls = DataLabelList(
            showVal=True,
            showCatName=False,
            showSerName=False,
            showLegendKey=False,
            numFmt=OVERLAY_LABEL_FORMAT,
            dLblPos=position,
            spPr=_label_graphical_properties(fill_color, border_color),
            txPr=_label_text_properties(text_color),
        )

    if cutoff_col is not None:
        cutoff = Reference(data_ws, min_col=cutoff_col, max_col=cutoff_col, min_row=first_row, max_row=last_row)
        chart.add_data(cutoff, titles_from_data=False)
        cutoff_series = chart.series[2]
        cutoff_series.tx = SeriesLabel(v="Cutoff")
        cutoff_series.graphicalProperties.line.noFill = True
        cutoff_series.marker.symbol = "none"
        cutoff_series.errBars = ErrorBars(
            errDir="y",
            errBarType="minus",
            errValType="fixedVal",
            noEndCap=True,
            val=1,
            spPr=GraphicalProperties(
                ln=LineProperties(solidFill=CUTOFF_RED, w=19050, prstDash="dash")
            ),
        )
        # Only the single non-#N/A cutoff point receives a label.  Category
        # name contributes the reporting date, so the visible tag reads
        # "Cutoff <date>" without worksheet shapes or VBA.
        cutoff_series.dLbls = DataLabelList(
            showVal=False,
            showCatName=True,
            showSerName=True,
            showLegendKey=False,
            dLblPos="t",
            separator=" ",
            spPr=_label_graphical_properties(CUTOFF_LABEL_BG, CUTOFF_LABEL_BORDER),
            txPr=_label_text_properties(CUTOFF_RED),
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
    """LW-12.4: independent-cutoff, project-bounded responsive overlays."""
    if not dataset.periods:
        return False, False
    if DATA_SHEET not in workbook.sheetnames:
        raise ValueError("Dashboard_Data is required before building traditional overlays.")
    data_ws = workbook[DATA_SHEET]

    scurve_plan_row = _scurve_plan_row(dataset)
    control_row = max(dataset.header_row + 1, scurve_plan_row - 1)
    bottom_anchor_row = max(OVERLAY_TOP_ROW + 1, scurve_plan_row - 1)

    weekly_list_last = max(2, len(dataset.periods) + 1)
    monthly_list_last = max(
        2,
        1 + sum(1 for r in range(2, data_ws.max_row + 1) if data_ws.cell(r, 11).value not in (None, "")),
    )

    dashboard_cutoff = workbook["Dashboard"]["K5"].value if "Dashboard" in workbook.sheetnames else None
    weekly_dates = [data_ws.cell(r, 10).value for r in range(2, weekly_list_last + 1) if data_ws.cell(r, 10).value not in (None, "")]
    monthly_dates = [data_ws.cell(r, 11).value for r in range(2, monthly_list_last + 1) if data_ws.cell(r, 11).value not in (None, "")]
    dashboard_day = _as_date(dashboard_cutoff)
    weekly_initial = next(
        (d for d in weekly_dates if _as_date(d) == dashboard_day),
        weekly_dates[-1] if weekly_dates else dashboard_cutoff,
    )
    eligible_monthly = [
        d for d in monthly_dates
        if dashboard_day is not None and _as_date(d) is not None and _as_date(d) <= dashboard_day
    ]
    monthly_initial = eligible_monthly[-1] if eligible_monthly else (monthly_dates[-1] if monthly_dates else dashboard_cutoff)

    weekly_cutoff_ref = None
    monthly_cutoff_ref = None
    if "main" in workbook.sheetnames:
        weekly_cutoff_ref = _add_cutoff_control(
            workbook["main"], dataset, row=control_row, list_col="J", list_last_row=weekly_list_last,
            initial_value=weekly_initial, display_format="dd/mm/yyyy",
        )
    if "main_monthly" in workbook.sheetnames:
        monthly_cutoff_ref = _add_cutoff_control(
            workbook["main_monthly"], dataset, row=control_row, list_col="K", list_last_row=monthly_list_last,
            initial_value=monthly_initial, display_format="mmm yyyy",
        )

    if weekly_cutoff_ref is None:
        weekly_cutoff_ref = "Dashboard!$K$5"
    if monthly_cutoff_ref is None:
        monthly_cutoff_ref = "Dashboard!$K$5"
    ensure_overlay_visible_actual_columns(
        workbook,
        weekly_cutoff_ref=weekly_cutoff_ref,
        monthly_cutoff_ref=monthly_cutoff_ref,
    )

    # Remove prior overlay charts when rebuilding the same workbook.
    for sheet_name in ("main", "main_monthly"):
        if sheet_name in workbook.sheetnames:
            workbook[sheet_name]._charts = []

    weekly_first, weekly_last, weekly_first_col, weekly_last_col = _weekly_project_window(dataset)
    monthly_first, monthly_last, monthly_first_col, monthly_last_col = _monthly_project_window(data_ws, dataset)

    weekly_added = False
    if "main" in workbook.sheetnames:
        chart = _overlay_chart(
            data_ws=data_ws,
            date_col=1,
            plan_col=2,
            actual_col=16,
            cutoff_col=18,
            first_row=weekly_first,
            last_row=weekly_last,
        )
        chart.anchor = _responsive_anchor(
            first_col=weekly_first_col,
            last_col=weekly_last_col,
            top_row=OVERLAY_TOP_ROW,
            bottom_row=bottom_anchor_row,
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
            cutoff_col=19,
            first_row=monthly_first,
            last_row=monthly_last,
        )
        chart.anchor = _responsive_anchor(
            first_col=monthly_first_col,
            last_col=monthly_last_col,
            top_row=OVERLAY_TOP_ROW,
            bottom_row=bottom_anchor_row,
        )
        workbook["main_monthly"].add_chart(chart)
        monthly_added = True

    return weekly_added, monthly_added
