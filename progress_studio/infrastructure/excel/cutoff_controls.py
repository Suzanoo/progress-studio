"""Excel-only placement and migration of the existing overlay controls.

Date calendars, defaults and chart calculation remain owned by the overlay
renderer. Placement never inserts rows or shifts user data/formula references.
"""
from copy import copy
from datetime import date, datetime

from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Protection
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

LABEL = "Cutoff Date"
CONTROL_NAMES = {
    "main": "PS_WEEKLY_OVERLAY_CUTOFF",
    "main_monthly": "PS_MONTHLY_OVERLAY_CUTOFF",
}


class CutoffPlacementError(ValueError):
    """A layout/content conflict prevents safe cutoff relocation."""


def generated_rows(dataset):
    rows = set()
    for item in dataset.rows:
        kind, pa = item.row_type.strip().lower(), item.pa.strip().upper()
        if kind in {"activity", "wbs", "project summary"} and pa == "P":
            rows.update((item.row_number, item.row_number + 1))
        elif kind == "s-curve" and pa in {"P", "AP", "A", "AA"}:
            rows.add(item.row_number)
    return rows


def footer_control_row(dataset):
    rows = generated_rows(dataset)
    if not rows or max(rows) >= 1048576:
        raise CutoffPlacementError("Cannot locate a safe generated footer for the cutoff control.")
    return max(rows) + 1


def bound_cutoff_cell(ws):
    """Resolve only a single local, labelled cell; malformed names fail closed."""
    definition = ws.parent.defined_names.get(CONTROL_NAMES.get(ws.title, ""))
    if definition is None:
        return None
    try:
        destinations = list(definition.destinations)
        if len(destinations) != 1 or destinations[0][0] != ws.title:
            return None
        area = CellRange(destinations[0][1])
        if area.size != {"columns": 1, "rows": 1} or area.min_col <= 1:
            return None
        cell = ws.cell(area.min_row, area.min_col)
        if not isinstance(cell, MergedCell) and ws.cell(cell.row, cell.column - 1).value == LABEL:
            return cell
    except (ValueError, TypeError, AttributeError):
        pass
    return None


def find_cutoff_controls(ws, dataset):
    """Find supported old/new pairs without confusing Activity descriptions.

    Old layouts used L:M or Description:Amount. WU-1 uses B:C. A labelled
    pair must be outside generated table rows and have a date/blank/linked value.
    """
    pairs = {(12, 13), (2, 3)}
    if dataset is not None:
        pairs.add((dataset.header_column("description") or 3,
                   dataset.header_column("amount") or 4))
    bound = bound_cutoff_cell(ws)
    if bound is not None:
        pairs.add((bound.column - 1, bound.column))
    occupied = generated_rows(dataset) if dataset is not None else set()
    controls = []
    for row in range((dataset.header_row if dataset else 4) + 1, ws.max_row + 1):
        if row in occupied:
            continue
        for label_col, value_col in sorted(pairs):
            if label_col == value_col or ws.cell(row, label_col).value != LABEL:
                continue
            cell = ws.cell(row, value_col)
            value = cell.value
            supported = value is None or isinstance(value, (date, datetime)) or value == "=PS_WEEKLY_OVERLAY_CUTOFF"
            if isinstance(value, str) and not supported:
                try:
                    datetime.fromisoformat(value)
                    supported = True
                except ValueError:
                    pass
            if not supported or isinstance(cell, MergedCell):
                raise CutoffPlacementError(f"Cannot relocate cutoff at '{ws.title}'!{cell.coordinate}: unsupported value/layout.")
            controls.append((row, label_col, value_col))
    return controls


def remove_cell_validation(ws, coordinate):
    """Subtract one cell, preserving validation on every other target."""
    for validation in list(ws.data_validations.dataValidation):
        if coordinate not in validation.cells:
            continue
        point = CellRange(coordinate)
        r, c = point.min_row, point.min_col
        remaining = []
        for area in validation.sqref.ranges:
            if coordinate not in area:
                remaining.append(str(area))
                continue
            for left, top, right, bottom in (
                (area.min_col, area.min_row, area.max_col, r - 1),
                (area.min_col, r + 1, area.max_col, area.max_row),
                (area.min_col, r, c - 1, r),
                (c + 1, r, area.max_col, r),
            ):
                if left <= right and top <= bottom:
                    remaining.append(str(CellRange(min_col=left, min_row=top, max_col=right, max_row=bottom)))
        if remaining:
            validation.sqref = MultiCellRange(" ".join(remaining))
        else:
            ws.data_validations.dataValidation.remove(validation)


def validate_placement(ws, row, controls):
    owned = {(r, col) for r, lc, vc in controls for col in (lc, vc)}
    for col in (2, 3):
        cell = ws.cell(row, col)
        if any(cell.coordinate in area for area in ws.merged_cells.ranges):
            raise CutoffPlacementError(f"Cutoff placement conflicts with merged cell '{ws.title}'!{cell.coordinate}.")
        if (row, col) not in owned and (
            cell.value is not None or cell.comment is not None or cell.hyperlink is not None
            or any(cell.coordinate in dv.cells for dv in ws.data_validations.dataValidation)
        ):
            raise CutoffPlacementError(f"Cutoff placement would overwrite existing content/control at '{ws.title}'!{cell.coordinate}. Move that content to a free cell and retry on a copy.")
    for r, lc, vc in controls:
        for col in (lc, vc):
            cell = ws.cell(r, col)
            if any(cell.coordinate in area for area in ws.merged_cells.ranges) or cell.hyperlink:
                raise CutoffPlacementError(f"Unsupported cutoff layout at '{ws.title}'!{cell.coordinate}.")


def clear_cutoff_controls(ws, controls):
    for row, label_col, value_col in controls:
        for col in (label_col, value_col):
            cell = ws.cell(row, col)
            remove_cell_validation(ws, cell.coordinate)
            cell.value = None
            cell.comment = None
            cell._style = copy(ws.cell(row, 1)._style)
            cell.protection = Protection(locked=True)


def prepare_cutoff_placement(ws, dataset):
    """Preflight before deleting old controls; return the stable footer row."""
    row = footer_control_row(dataset)
    controls = find_cutoff_controls(ws, dataset)
    validate_placement(ws, row, controls)
    clear_cutoff_controls(ws, controls)
    return row


def validate_existing_cutoff_placements(workbook):
    """Check original controls before regeneration discards generated views."""
    from progress_studio.infrastructure.excel.main_dataset_workbook_adapter import main_dataset_from_workbook
    dataset = main_dataset_from_workbook(workbook)
    row = footer_control_row(dataset)
    for name in CONTROL_NAMES:
        if name in workbook.sheetnames:
            ws = workbook[name]
            validate_placement(ws, row, find_cutoff_controls(ws, dataset))
