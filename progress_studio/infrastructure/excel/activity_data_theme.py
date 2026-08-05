from __future__ import annotations

from copy import copy
from dataclasses import dataclass

from openpyxl.styles import Border, Font, PatternFill, Side


@dataclass(frozen=True)
class ActivityDataPalette:
    """Colors used only by the Activity Data section of the main sheet."""

    wbs_level_1_fill: str = "F4B183"
    wbs_level_2_fill: str = "F8CBAD"
    font_color: str = "000000"
    separator_color: str = "C65911"


DEFAULT_ACTIVITY_DATA_PALETTE = ActivityDataPalette()


def _normalized(value: object) -> str:
    return str(value or "").strip().lower()


def _as_level(value: object) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _activity_data_last_column(ws, header_row: int) -> int:
    """Return the column before the first weekly timescale date.

    The weekly dates are stored on the same row as the Activity Data headers.
    Keeping this boundary explicit guarantees that this formatter never changes
    timescale fills, borders, or bars.
    """
    for col in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col).value
        if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
            return col - 1
    return ws.max_column


def apply_activity_data_wbs_hierarchy(
    ws,
    *,
    header_row: int = 4,
    first_data_row: int = 5,
    palette: ActivityDataPalette = DEFAULT_ACTIVITY_DATA_PALETTE,
) -> None:
    """Differentiate WBS level 1 and 2 in Activity Data only.

    Project Summary, Activity rows, and all timescale cells are intentionally
    left unchanged. Actual rows inherit the style of their paired WBS Plan row.
    """
    headers = {
        _normalized(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if _normalized(ws.cell(header_row, col).value)
    }
    row_type_col = headers.get("row type")
    outline_col = headers.get("outline level")
    pa_col = headers.get("p/a")
    if row_type_col is None or outline_col is None:
        return

    last_data_col = _activity_data_last_column(ws, header_row)
    level_1_fill = PatternFill("solid", fgColor=palette.wbs_level_1_fill)
    level_2_fill = PatternFill("solid", fgColor=palette.wbs_level_2_fill)
    medium_top = Side(style="medium", color=palette.separator_color)

    previous_wbs_level: int | None = None
    for row in range(first_data_row, ws.max_row + 1):
        row_type = _normalized(ws.cell(row, row_type_col).value)
        pa = str(ws.cell(row, pa_col).value or "").strip().upper() if pa_col else ""

        if row_type == "wbs":
            level = _as_level(ws.cell(row, outline_col).value)
            previous_wbs_level = level
        elif pa == "A" and previous_wbs_level in (1, 2):
            level = previous_wbs_level
        else:
            previous_wbs_level = None
            continue

        if level not in (1, 2):
            continue
        fill = level_1_fill if level == 1 else level_2_fill

        for col in range(1, last_data_col + 1):
            cell = ws.cell(row, col)
            cell.fill = copy(fill)
            cell.font = copy(cell.font)
            cell.font = Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=True,
                italic=cell.font.italic,
                vertAlign=cell.font.vertAlign,
                underline=cell.font.underline,
                strike=cell.font.strike,
                color=palette.font_color,
            )
            if level == 1:
                old = cell.border
                cell.border = Border(
                    left=copy(old.left),
                    right=copy(old.right),
                    top=medium_top,
                    bottom=copy(old.bottom),
                    diagonal=copy(old.diagonal),
                    diagonal_direction=old.diagonal_direction,
                    diagonalUp=old.diagonalUp,
                    diagonalDown=old.diagonalDown,
                    outline=old.outline,
                    vertical=copy(old.vertical),
                    horizontal=copy(old.horizontal),
                )
