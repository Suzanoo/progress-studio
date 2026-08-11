
from __future__ import annotations

from copy import copy
from datetime import date

from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from progress_studio.domain.main_dataset import MainDataset
from progress_studio.domain.monthly_cache import MonthlyCache
from progress_studio.infrastructure.excel.timescale_workbook import (
    DATE_FILL,
    HEADER_BORDER,
    MONTH_FILL,
    WEEK_FILL,
    YEAR_FILL,
)
from progress_studio.infrastructure.excel.progress_workbook import (
    add_progress_conditional_formatting,
    clear_progress_conditional_formatting,
    clear_timescale_direct_fills,
    SCURVE_PLAN_FILL,
    SCURVE_ACTUAL_FILL,
    WBS_PLAN_FILL,
    WBS_ACTUAL_FILL,
)


def build_live_monthly_view(
    workbook,
    dataset: MainDataset,
    cache: MonthlyCache,
    *,
    source_sheet: str = "main",
    target_sheet: str = "main_monthly",
) -> int:
    """LW-10.0 Full Live Monthly baseline.

    Every monthly timescale cell links directly to the corresponding weekly range
    in `main`. This intentionally maximizes live behavior before later LW-10.x
    milestones cut formula complexity/volume based on measured pain points.
    """
    if source_sheet not in workbook.sheetnames:
        raise ValueError(f"Monthly source worksheet was not found: {source_sheet}")
    if not dataset.periods:
        raise ValueError("Weekly periods were not found in MainDataset.")

    source = workbook[source_sheet]
    if target_sheet in workbook.sheetnames:
        del workbook[target_sheet]

    monthly = workbook.copy_worksheet(source)
    monthly.title = target_sheet
    source_index = workbook.worksheets.index(source)
    workbook._sheets.remove(monthly)
    workbook._sheets.insert(source_index + 1, monthly)
    monthly.cell(1, 1).value = "Activity Data — Monthly View"

    first_timescale_col = dataset.periods[0].column

    for merged in list(monthly.merged_cells.ranges):
        if merged.max_col >= first_timescale_col:
            monthly.unmerge_cells(str(merged))

    template_styles = {
        row: copy(monthly.cell(row, first_timescale_col)._style)
        for row in range(5, monthly.max_row + 1)
    }
    template_formats = {
        row: monthly.cell(row, first_timescale_col).number_format
        for row in range(5, monthly.max_row + 1)
    }

    monthly.delete_cols(
        first_timescale_col,
        monthly.max_column - first_timescale_col + 1,
    )

    # Year/month/date grammar.
    current_year = None
    year_start = None
    for index, period in enumerate(cache.periods, start=1):
        col = first_timescale_col + index - 1
        reporting = period.reporting_date
        month_name = reporting.strftime("%B") if reporting else f"Month {index}"
        year = reporting.year if reporting else None

        monthly.cell(2, col).value = month_name
        monthly.cell(2, col).fill = MONTH_FILL
        monthly.cell(2, col).font = Font(color="000000", bold=True)
        monthly.cell(2, col).alignment = Alignment(horizontal="center", vertical="center")
        monthly.cell(2, col).border = HEADER_BORDER

        monthly.cell(3, col).value = period.key
        monthly.cell(3, col).fill = WEEK_FILL
        monthly.cell(3, col).font = Font(color="000000", bold=True)
        monthly.cell(3, col).alignment = Alignment(horizontal="center", vertical="center")
        monthly.cell(3, col).border = HEADER_BORDER

        monthly.cell(4, col).value = reporting
        monthly.cell(4, col).fill = DATE_FILL
        monthly.cell(4, col).font = Font(color="000000", bold=True)
        monthly.cell(4, col).alignment = Alignment(horizontal="center", vertical="center")
        monthly.cell(4, col).border = HEADER_BORDER
        monthly.cell(4, col).number_format = "dd/mm/yy"
        monthly.column_dimensions[get_column_letter(col)].width = 12

        monthly.cell(1, col).value = str(year) if year is not None else ""
        monthly.cell(1, col).fill = YEAR_FILL
        monthly.cell(1, col).font = Font(color="FFFFFF", bold=True)
        monthly.cell(1, col).alignment = Alignment(horizontal="center", vertical="center")
        monthly.cell(1, col).border = HEADER_BORDER

    source_ref = "'" + source_sheet.replace("'", "''") + "'"
    by_source_row = {row.source_row: row for row in cache.rows}
    for source_row, cached in by_source_row.items():
        if source_row > monthly.max_row:
            continue
        for index, period in enumerate(cache.periods, start=0):
            col = first_timescale_col + index
            cell = monthly.cell(source_row, col)
            if source_row in template_styles:
                cell._style = copy(template_styles[source_row])
                cell.number_format = template_formats[source_row]

            first_week = get_column_letter(period.source_columns[0])
            last_week = get_column_letter(period.source_columns[-1])
            source_range = (
                f"{source_ref}!{first_week}{source_row}:{last_week}{source_row}"
            )

            row_type = cached.row_type.strip().lower()
            pa = cached.pa.strip().upper()
            monthly_date_ref = f"{get_column_letter(col)}$4"

            if row_type == "s-curve":
                # Monthly S-Curve is a true live view of main.
                # P/AP remain full baseline. A/AA stop at the Dashboard cutoff.
                if pa == "AP":
                    cell.value = f"={source_ref}!{last_week}{source_row}"
                elif pa == "AA":
                    cell.value = (
                        f'=IF({monthly_date_ref}>Dashboard!$K$5,"",'
                        f'{source_ref}!{last_week}{source_row})'
                    )
                elif pa == "A":
                    cell.value = (
                        f'=IF({monthly_date_ref}>Dashboard!$K$5,"",'
                        f'IF(COUNT({source_range})=0,"",SUM({source_range})))'
                    )
                else:  # S-Curve Plan
                    cell.value = f'=IF(COUNT({source_range})=0,"",SUM({source_range}))'
            else:
                # Full-live baseline intentionally uses the same straightforward
                # formula for Project/WBS/Activity Plan and Actual rows.
                cell.value = f'=IF(COUNT({source_range})=0,"",SUM({source_range}))'

    # Match main exactly: blank timescale cells have no fill; populated cells
    # are colored by the same Project/WBS/Activity Plan/Actual CF rules.
    monthly_timescale_cols = list(
        range(first_timescale_col, first_timescale_col + len(cache.periods))
    )
    header_columns = {name: col for name, col in dataset.headers}
    required = ("row type", "activity id", "p/a", "outline level")
    if monthly_timescale_cols and all(name in header_columns for name in required):
        clear_timescale_direct_fills(
            monthly,
            monthly_timescale_cols,
            monthly.max_row,
        )
        clear_progress_conditional_formatting(monthly, monthly_timescale_cols)
        add_progress_conditional_formatting(
            monthly,
            monthly_timescale_cols,
            header_columns["row type"],
            header_columns["activity id"],
            header_columns["p/a"],
            header_columns["outline level"],
            monthly.max_row,
        )

        # S-Curve summary rows sit outside the Project/WBS/Activity CF grammar.
        # Paint their timescale explicitly with the same palette used by main.
        pa_col = header_columns["p/a"]
        row_type_col = header_columns["row type"]
        scurve_fills = {
            "P": SCURVE_PLAN_FILL,
            "AP": WBS_PLAN_FILL,
            "A": SCURVE_ACTUAL_FILL,
            "AA": WBS_ACTUAL_FILL,
        }
        for row in range(dataset.header_row + 1, monthly.max_row + 1):
            row_type = str(monthly.cell(row, row_type_col).value or "").strip().lower()
            if row_type != "s-curve":
                continue
            pa = str(monthly.cell(row, pa_col).value or "").strip().upper()
            fill = scurve_fills.get(pa)
            if fill is None:
                continue
            for col in monthly_timescale_cols:
                monthly.cell(row, col).fill = copy(fill)

    try:
        monthly.data_validations.dataValidation = []
    except AttributeError:
        pass
    monthly.freeze_panes = monthly.cell(5, first_timescale_col)
    monthly.sheet_properties.outlinePr.summaryBelow = False
    monthly.sheet_properties.outlinePr.applyStyles = True
    monthly.sheet_view.showGridLines = source.sheet_view.showGridLines
    return len(cache.periods)
