"""Strict read-only monetary evidence layered on the sparse XLSX reader.

Never save a workbook. Do not use MainDataset's float coercion or projection:
those intentionally discard evidence needed by the writeback boundary.
"""
from io import BytesIO
import re
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from progress_studio.domain.xml_writeback import WorkbookAmount, WorkbookAmounts, monetary_decimal
from .rebuild_workbook_reader import RebuildWorkbookReader as Reader, MAIN_NS, REL_NS, PKG_REL_NS

NS = '{' + MAIN_NS + '}'
MILESTONE_PREFIX = 'Calendar Duration Milestone: '


def _rows(root):
    return root.findall(f'{NS}sheetData/{NS}row')


def _cells(row):
    number = row.get('r', '')
    if not re.fullmatch(r'[1-9][0-9]*', number):
        raise ValueError('Invalid worksheet row address.')
    for cell in row.findall(NS + 'c'):
        ref = re.fullmatch(r'[A-Z]+([1-9][0-9]*)', cell.get('r', ''))
        if ref is None or ref.group(1) != number:
            raise ValueError('Invalid worksheet cell/row address.')
    cells = Reader._row_cells(row)
    if len(cells) != len(row.findall(NS + 'c')):
        raise ValueError('Duplicate or invalid worksheet cell addresses.')
    return cells


def _text(cell, shared):
    return Reader._cell_text(cell, shared).strip()


def _literal(cell, shared):
    if cell is not None and (cell.find(NS + 'f') is not None or cell.get('t') == 'e'):
        raise ValueError('Activity identity/classification must be a literal, not a formula/error.')
    return _text(cell, shared)


def _header(rows, shared, required):
    def key(text):
        return ''.join(text.lower().split())
    for row in rows:
        if int(row.get('r', '0')) > 30:
            continue
        found = {}
        duplicates = set()
        for col, cell in _cells(row).items():
            label = key(_text(cell, shared))
            if label in found:
                duplicates.add(label)
            found[label] = col
        if set(required).issubset(found):
            if duplicates.intersection(required):
                raise ValueError('Duplicate required worksheet headers.')
            for label in required:
                _literal(_cells(row)[found[label]], shared)
            return int(row.get('r')), found
    raise ValueError('Missing worksheet headers: ' + ', '.join(required))


def _flag(cell, shared):
    raw = _literal(cell, shared)
    if cell is None or cell.get('t', 'n') not in ('n', 'b') or raw not in ('0', '1'):
        raise ValueError('Milestone flag must be an explicit TRUE/FALSE (0/1).')
    return raw == '1'


def read_workbook_amounts(data: bytes) -> WorkbookAmounts:
    errors, result = [], []
    try:
        with ZipFile(BytesIO(data)) as package:
            if len(package.namelist()) != len(set(package.namelist())):
                raise ValueError('Duplicate XLSX package entries.')
            workbook = ET.fromstring(package.read('xl/workbook.xml'))
            relationships = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
            relationship_ids = [r.get('Id') for r in relationships.findall('{' + PKG_REL_NS + '}Relationship')]
            if None in relationship_ids or len(relationship_ids) != len(set(relationship_ids)):
                raise ValueError('Ambiguous workbook relationship identities.')
            sheets = workbook.findall(f'{NS}sheets/{NS}sheet')
            names = [s.get('name', '').casefold() for s in sheets]
            relids = [s.get('{' + REL_NS + '}id') for s in sheets]
            if len(names) != len(set(names)) or len(relids) != len(set(relids)):
                raise ValueError('Ambiguous workbook sheet identities.')
            sheet_map = Reader._sheet_map(package)
            shared = Reader._shared_strings(package)
            rows = _rows(ET.fromstring(package.read(sheet_map['main'])))
            numbers = [r.get('r') for r in rows]
            if len(numbers) != len(set(numbers)):
                raise ValueError('Duplicate worksheet row numbers.')
            header, columns = _header(rows, shared, ('rowtype', 'p/a', 'activityid', 'amount'))
            # WU-2 stores positive milestone markers in Info. Missing markers
            # do not certify ordinary identity. EV flags supply explicit true/false.
            flags = {}
            info = {}
            if 'Info' in sheet_map:
                for row in _rows(ET.fromstring(package.read(sheet_map['Info']))):
                    cells = _cells(row)
                    key = _text(cells.get(1), shared)
                    if key.startswith(MILESTONE_PREFIX):
                        aid = key[len(MILESTONE_PREFIX):].strip().casefold()
                        if not aid or aid in info:
                            raise ValueError('Duplicate/blank Info milestone identity.')
                        info[aid] = _flag(cells.get(2), shared)
            if 'EV Monetary Inputs' in sheet_map:
                evrows = _rows(ET.fromstring(package.read(sheet_map['EV Monetary Inputs'])))
                evheader, evcols = _header(evrows, shared, ('activityid', 'milestone'))
                for row in evrows:
                    if int(row.get('r', '0')) <= evheader:
                        continue
                    cells = _cells(row)
                    aid = _literal(cells.get(evcols['activityid']), shared)
                    if not aid:
                        if _text(cells.get(evcols['milestone']), shared):
                            raise ValueError('Blank EV milestone Activity ID.')
                        continue
                    key = aid.casefold()
                    if key in flags:
                        raise ValueError('Duplicate/case-colliding EV milestone Activity ID: ' + aid)
                    flags[key] = _flag(cells.get(evcols['milestone']), shared)
            for key, flag in info.items():
                if key in flags and flags[key] != flag:
                    errors.append(f'{key}: conflicting Info/EV milestone classification.')
                flags[key] = flag
            seen = set()
            for row in rows:
                number = int(row.get('r', '0'))
                if number <= header:
                    continue
                cells = _cells(row)
                def cell(label):
                    return cells.get(columns.get(label))
                try:
                    kind = _literal(cell('rowtype'), shared).lower()
                    pa = _literal(cell('p/a'), shared).upper()
                    if kind != 'activity':
                        if _text(cell('activityid'), shared) and pa == 'P':
                            raise ValueError('Activity ID on a non-Activity Plan row.')
                        continue
                    if pa not in ('P', 'A'):
                        raise ValueError('Activity row needs P or A classification.')
                    if pa == 'A':
                        continue
                    aid = _literal(cell('activityid'), shared)
                    if not aid:
                        raise ValueError('Blank Activity ID on Plan row.')
                    if aid.casefold() in seen:
                        raise ValueError('Duplicate/case-colliding Activity ID: ' + aid)
                    seen.add(aid.casefold())
                    amount_cell = cell('amount')
                    raw = _text(amount_cell, shared)
                    formula = amount_cell is not None and amount_cell.find(NS + 'f') is not None
                    value = None
                    try:
                        if amount_cell is None or amount_cell.get('t', 'n') != 'n':
                            raise ValueError('Amount must be a numeric cell or saved numeric formula result.')
                        value = monetary_decimal(raw)
                    except ValueError as exc:
                        errors.append(f'main row {number}, {aid}: {exc}' + (' (formula cache)' if formula else ''))
                    result.append(WorkbookAmount(aid, _text(cell('description'), shared), number,
                                                 raw, value, formula, flags.get(aid.casefold())))
                except ValueError as exc:
                    errors.append(f'main row {number}: {exc}')
    except (ValueError, KeyError, OSError, BadZipFile, ET.ParseError) as exc:
        errors.append(f'Workbook structure: {exc}')
    if not result:
        errors.append('No Activity Plan rows available for writeback.')
    return WorkbookAmounts(tuple(result), tuple(errors))
