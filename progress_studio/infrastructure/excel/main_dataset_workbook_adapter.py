from __future__ import annotations

import re
from datetime import date, datetime

from progress_studio.domain.main_dataset import MainDataset, MainPeriod, MainRow


def _as_date(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    return None


def _as_float(value):
    if value in (None, "") or isinstance(value, str) and value.startswith("="):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_int(value):
    number = _as_float(value)
    return int(number) if number is not None else None


def main_dataset_from_workbook(workbook, *, workbook_name: str = "in-memory.xlsx") -> MainDataset:
    """Read the already-open `main` worksheet into MainDataset without reopening XLSX."""
    ws = workbook["main"]
    header_row = None
    headers: dict[str, int] = {}
    for row in range(1, min(ws.max_row, 30) + 1):
        found = {
            str(ws.cell(row, col).value or "").strip().lower(): col
            for col in range(1, ws.max_column + 1)
            if str(ws.cell(row, col).value or "").strip()
        }
        if {"row type", "p/a", "activity id"}.issubset(found):
            header_row, headers = row, found
            break
    if header_row is None:
        raise ValueError("Worksheet 'main' is missing Row Type / P/A / Activity ID headers.")

    periods: list[MainPeriod] = []
    for col in range(1, ws.max_column + 1):
        key = str(ws.cell(header_row - 1, col).value or "").strip()
        if re.fullmatch(r"W\d+", key, flags=re.IGNORECASE):
            periods.append(MainPeriod(col, key, _as_date(ws.cell(header_row, col).value)))

    rows: list[MainRow] = []
    for row in range(header_row + 1, ws.max_row + 1):
        def val(name: str):
            col = headers.get(name)
            return ws.cell(row, col).value if col else None
        rows.append(MainRow(
            row_number=row,
            row_type=str(val("row type") or "").strip(),
            pa=str(val("p/a") or "").strip(),
            wbs=str(val("wbs") or "").strip(),
            description=str(val("description") or "").strip(),
            activity_id=str(val("activity id") or "").strip(),
            outline_level=_as_int(val("outline level")),
            plan_start=_as_date(val("plan start")),
            plan_finish=_as_date(val("plan finish")),
            amount=_as_float(val("amount")),
            percent_complete=_as_float(val("% complete")),
            period_values=tuple((p.column, _as_float(ws.cell(row, p.column).value)) for p in periods),
        ))
    return MainDataset(
        workbook_name=workbook_name,
        header_row=header_row,
        headers=tuple(sorted(headers.items(), key=lambda item: item[1])),
        periods=tuple(periods),
        rows=tuple(rows),
    )
