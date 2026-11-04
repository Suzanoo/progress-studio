"""Strict source evidence for Amount-only updates; ordinary BOQ import is unchanged."""
from dataclasses import dataclass
import math
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook

from progress_studio.domain.mapping_models import BOQRow
from progress_studio.infrastructure.excel.mapping_reader import BOQSheetReader
from progress_studio.infrastructure.excel.rebuild_workbook_reader import RebuildWorkbookReader, MAIN_NS


@dataclass(frozen=True)
class BOQRevision:
    # Every non-Amount cell, including excluded/zero rows, participates in identity.
    structure: tuple
    amounts: tuple[tuple[int, float | None], ...]
    formula_rows: tuple[int, ...]
    items: tuple[BOQRow, ...]


def read_boq_revision(path: Path, sheet_name: str) -> BOQRevision:
    # Excel writes an intentional empty-string formula result as t="str", <v/>.
    # Distinguish unchanged BOQ headings from an absent numeric formula cache.
    # Never infer a result by evaluating or following a formula reference.
    with ZipFile(path) as package:
        sheets = RebuildWorkbookReader._sheet_map(package)
        if sheet_name not in sheets:
            raise ValueError(f"BOQ worksheet is missing: {sheet_name}")
        root = ET.fromstring(package.read(sheets[sheet_name]))
        ns = '{' + MAIN_NS + '}'
        empty_results = {c.get('r') for c in root.iter(ns+'c')
                         if c.get('t') == 'str' and c.find(ns+'f') is not None
                         and c.find(ns+'v') is not None and not c.find(ns+'v').text}
    formulas = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        values = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            if sheet_name not in formulas.sheetnames:
                raise ValueError(f"BOQ worksheet is missing: {sheet_name}")
            sheet, cached = formulas[sheet_name], values[sheet_name]
            header = next(sheet.iter_rows(min_row=1, max_row=1))
            names = [str(c.value or '').strip().lower() for c in header]
            if names.count('amount') != 1:
                raise ValueError('BOQ must contain exactly one Amount column.')
            for name in ('wbs-2', 'wbs-3', 'wbs-4', 'description'):
                if names.count(name) != 1:
                    raise ValueError(f'BOQ must contain exactly one {name} column.')
            amount_col = names.index('amount')
            structure, amounts, formula_rows = [], [], []
            identities = set()
            for row_no, (row, saved) in enumerate(zip(sheet.iter_rows(), cached.iter_rows()), 1):
                if not any(c.value is not None for c in row) and not any(c.value is not None for c in saved):
                    continue
                for i, cell in enumerate(row):
                    if i != amount_col and cell.data_type == 'f':
                        result = saved[i]
                        if result.data_type == 'e' or (result.value is None and cell.coordinate not in empty_results):
                            raise ValueError(f'{sheet_name}!{cell.coordinate}: non-Amount identity formula has no usable saved result. Recalculate and Save in Excel.')
                # Position is only one part of evidence: compare all non-Amount
                # source content AND saved values, not a row number/name match.
                identity = tuple((i, c.data_type, c.value, saved[i].data_type, saved[i].value)
                                 for i, c in enumerate(row)
                                 if (row_no == 1 or i != amount_col)
                                 and (c.value is not None or saved[i].value is not None))
                structure.append((row_no, identity))
                if row_no == 1:
                    continue
                cell, result = row[amount_col], saved[amount_col]
                value = result.value
                if cell.data_type == 'f':
                    formula_rows.append(row_no)
                    if value is None and cell.coordinate not in empty_results:
                        raise ValueError(f'{sheet_name}!{cell.coordinate}: missing formula result. Recalculate and Save in Excel.')
                if value is not None:
                    if (result.data_type == 'e' or isinstance(value, bool)
                            or not isinstance(value, (int, float))
                            or not math.isfinite(value) or value < 0):
                        raise ValueError(f'{sheet_name}!{cell.coordinate}: Amount must be a finite, non-negative saved number.')
                    value = float(value)
                amounts.append((row_no, value))
                if value is not None and value > 0:
                    if not identity or identity in identities:
                        raise ValueError(f'{sheet_name} row {row_no}: ambiguous duplicate BOQ identity; update rejected.')
                    identities.add(identity)
            items = tuple(BOQSheetReader().read(path, sheet_name))
            if len({item.key for item in items}) != len(items):
                raise ValueError('Duplicate BOQ keys; update rejected.')
            return BOQRevision(tuple(structure), tuple(amounts), tuple(formula_rows), items)
        finally:
            values.close()
    finally:
        formulas.close()
