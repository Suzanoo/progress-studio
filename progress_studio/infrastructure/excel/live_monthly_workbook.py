
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


def build_live_monthly_view(
    workbook,
    dataset: MainDataset,
    cache: MonthlyCache,
    *,
    source_sheet: str = "main",
    target_sheet: str = "main_monthly",
) -> int:
    """Render cached monthly values with one mutable workbook pass.

    The monthly timescale contains values only. No monthly cell references weekly
    cells in `main`, so the Live output avoids a monthly formula dependency graph.
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

    by_source_row = {row.source_row: row for row in cache.rows}
    for source_row, cached in by_source_row.items():
        if source_row > monthly.max_row:
            continue
        for index, value in enumerate(cached.values, start=0):
            col = first_timescale_col + index
            cell = monthly.cell(source_row, col)
            if source_row in template_styles:
                cell._style = copy(template_styles[source_row])
                cell.number_format = template_formats[source_row]
            cell.value = value if value is not None else ""

    try:
        monthly.data_validations.dataValidation = []
    except AttributeError:
        pass
    monthly.freeze_panes = monthly.cell(5, first_timescale_col)
    monthly.sheet_properties.outlinePr.summaryBelow = False
    monthly.sheet_properties.outlinePr.applyStyles = True
    monthly.sheet_view.showGridLines = source.sheet_view.showGridLines
    return len(cache.periods)
