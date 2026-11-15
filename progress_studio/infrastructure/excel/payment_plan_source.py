"""Build a report from one byte snapshot of an embedded Payment workbook."""
from dataclasses import replace
from decimal import Decimal, InvalidOperation, localcontext
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from urllib.parse import quote
from xml.etree import ElementTree as ET
from zipfile import ZipFile, BadZipFile

from progress_studio.domain.payment_plan_report import Leaf, Period, ReportSource, Row
from progress_studio.domain.payment_requirement_states import requirement_states
from progress_studio.infrastructure.excel.payment_input_reader import PaymentInputSparseReader
from progress_studio.infrastructure.excel.payment_input_summary import project_summary_tree
from progress_studio.infrastructure.excel.payment_progress_index import ActivityProgressIndexReader
from progress_studio.infrastructure.excel.payment_workbook import (
    MAIN_NS, PaymentWorkbookError, PaymentWorkbookSnapshotter,
)
from progress_studio.services.payment_position_engine import PaymentPositionEngine

NS = '{' + MAIN_NS + '}'
READER = PaymentInputSparseReader()


def literal(cell, strings):
    if cell is not None and cell.find(NS + 'f') is not None:
        raise PaymentWorkbookError(f"Formula at {cell.get('r')}. Prepare literal source values and save.")
    return READER._cell_text(cell, strings).strip()


def amount(cell, strings):
    if cell is not None and cell.find(NS + 'f') is not None:
        literal(cell, strings)
    if cell is None or cell.get('t', 'n') != 'n':
        return None
    try:
        value = Decimal(literal(cell, strings))
    except InvalidOperation:
        return None
    return value if value.is_finite() and value >= 0 else None


def headers(row, strings, required):
    found = {}
    for index, cell in READER._row_cells(row).items():
        name = READER._cell_text(cell, strings).strip().lower()
        if name in required:
            if name in found:
                raise PaymentWorkbookError(f'Duplicate source column {name}.')
            found[name] = index
    if missing := required - found.keys():
        raise PaymentWorkbookError(f'Missing {sorted(missing)}. Prepare embedded Payment Input.')
    return found


def monetary_origin(info, sheets):
    if 'BOQ Activity Mapping' in sheets:
        return 'BOQ Activity Mapping'
    if (info.get('Creation Weight Basis') == 'amount'
            and info.get('Amount Initialization Source', 'xml_field') == 'xml_field'
            and info.get('Creation Amount Field identity')):
        return 'XML Amount / ' + info['Creation Amount Field identity']
    raise PaymentWorkbookError('Unverified monetary basis. Use a monetary XML or BOQ-mapped source.')


def main_plan(main, strings, progress):
    rows = main.findall(f'{NS}sheetData/{NS}row')
    header, _ = ActivityProgressIndexReader()._find_headers(rows, strings)
    required = {'row type', 'p/a', 'activity id', 'wbs', 'description', 'outline level', 'amount'}
    labels = headers(next(r for r in rows if int(r.get('r')) == header), strings, required)
    header_cells = READER._row_cells(next(r for r in rows if int(r.get('r')) == header))
    for column, _, _ in progress.timescale_columns:
        literal(header_cells[column], strings)
    records = []
    unreliable = set()
    for row in rows:
        cells = READER._row_cells(row)
        def field(name):
            return literal(cells.get(labels[name]), strings)
        if field('p/a') != 'P' or field('row type') not in ('Project Summary', 'Activity', 'WBS'):
            continue
        activity = field('row type') == 'Activity'
        aid = field('activity id')
        if activity:
            for col, _, _ in progress.timescale_columns:
                c = cells.get(col)
                if c is not None and (c.find(NS+'f') is not None or
                                      (literal(c, strings) and amount(c, strings) is None)):
                    unreliable.add(aid)
                    break
        kind = 'ACT' if activity else ('PROJECT' if field('row type') == 'Project Summary' else 'WBS')
        records.append(dict(row_type=kind, activity_id=aid, wbs=field('wbs'),
                            activity_name=field('description'),
                            outline_level=int(field('outline level') or 0),
                            amount_cell=cells.get(labels['amount'])))
    return project_summary_tree(records), unreliable


def source_tree(rows, labels, main, strings, periods):
    leaves = {}
    tree = []
    members = {}
    ancestors = []
    issues = []
    active = [(row, READER._row_cells(row)) for row in rows if int(row.get('r')) >= 8]
    active = [(row, cells) for row, cells in active if any(
        READER._cell_text(cells.get(labels[n]), strings).strip()
        for n in ('type', 'activity id', 'activity name', 'wbs'))]
    if len(active) != len(main):
        raise PaymentWorkbookError('Schedule rows changed. Prepare/Reconcile Payment Input.')
    for (row, cells), item in zip(active, main):
        activity = item['row_type'] == 'ACT'
        aid, code = item['activity_id'], item.get('wbs', '')
        name, level = item['activity_name'], item['outline_level']
        index = int(row.get('r'))
        def field(label):
            return literal(cells.get(labels[label]), strings)
        if (field('type'), field('activity id'), field('wbs'), field('activity name')) != (
                'ACT' if activity else 'WBS', aid, code, name):
            raise PaymentWorkbookError(f'Hierarchy or identity differs at row {index}. Prepare/Reconcile.')
        outline = row.get('outlineLevel')
        if outline is not None and int(outline) != min(level, 7):
            raise PaymentWorkbookError(f'Stale Payment Input outline at row {index}. Prepare/Reconcile.')
        while ancestors and ancestors[-1][0] >= level:
            ancestors.pop()
        parent = ancestors[-1][1] if ancestors else None
        if not tree and (activity or level != 0):
            raise PaymentWorkbookError('Missing project summary. Prepare Payment Input.')
        if not ancestors and tree:
            raise PaymentWorkbookError('Missing common project root. Prepare Payment Input.')
        if activity:
            key = 'ACT/' + aid
            snapshot = amount(cells.get(labels['amount']), strings)
            current = amount(item.get('amount_cell'), strings)
            if snapshot != current:
                raise PaymentWorkbookError(f'Stale Amount for {aid}. Prepare/Reconcile and save.')
            if snapshot is None:
                issues.append(f'AmountUnavailable: {aid} (row {index}). Inspect main and Prepare.')
            entries = {}
            for p in periods:
                c = cells.get(p.column_index)
                t = literal(c, strings)
                if t:
                    if c.get('t') in ('e', 'b'):
                        raise PaymentWorkbookError(f"Invalid requirement at {c.get('r')}.")
                    if t.endswith('%'):
                        value = Decimal(t[:-1].strip())
                        with localcontext() as context:
                            context.prec = max(len(value.as_tuple().digits), 28)
                            v = value.scaleb(-2)
                    else:
                        v = Decimal(t)
                    if not v.is_finite() or not 0 <= v <= 1:
                        raise PaymentWorkbookError(f"Invalid requirement at {c.get('r')}.")
                    entries[p.period_id] = v, index, p.column_index
            states = tuple(requirement_states((p.period_id for p in periods), entries))
            leaves[aid] = Leaf(snapshot, states)
            for _, ancestor in ancestors:
                members[ancestor].append(aid)
            ids = [aid]
        else:
            key = 'PROJECT' if parent is None else parent + '/WBS/' + quote(code, safe='')
            if parent is not None and not code:
                raise PaymentWorkbookError(f'Missing WBS code at row {index}.')
            ancestors.append((level, key))
            ids = []
        if key in members:
            raise PaymentWorkbookError(f'Ambiguous reporting identity {key}.')
        members[key] = ids
        tree.append((key, parent, code, name, index, activity))
    return tuple(Row(k, p, c, n, i, tuple(members[k]), a) for k,p,c,n,i,a in tree), leaves, issues


def read_payment_plan_source(workbook, *, expected_revision=None):
    path = Path(workbook)
    data = path.read_bytes()
    revision = sha256(data).hexdigest()
    if expected_revision is not None and expected_revision != revision:
        raise PaymentWorkbookError('Selected workbook revision changed. Reload and review.')
    try:
        with ZipFile(BytesIO(data)) as z:
            sheets = PaymentWorkbookSnapshotter._sheet_map(z)
            if not {'main', 'Payment Input', 'Info'} <= sheets.keys():
                raise PaymentWorkbookError('Select the embedded Payment source with main, Info and Payment Input.')
            strings = PaymentWorkbookSnapshotter._shared_strings(z)
            def sheet(name):
                return ET.fromstring(z.read(sheets[name]))
            info = {}
            for row in sheet('Info').iter(NS+'row'):
                cells = READER._row_cells(row)
                key = READER._cell_text(cells.get(1), strings)
                if key in ('Project','Creation Weight Basis','Amount Initialization Source','Creation Amount Field identity'):
                    if key in info:
                        raise PaymentWorkbookError(f'Duplicate Info field {key}.')
                    info[key] = literal(cells.get(2), strings)
            origin = monetary_origin(info, sheets)
            if not info.get('Project'):
                raise PaymentWorkbookError('Missing project identity in Info.')
            main = sheet('main')
            source = sheet('Payment Input')
            payment = READER._parse(path, source, strings)
            progress = ActivityProgressIndexReader()._parse(path, main, strings)
            plan, unreliable = main_plan(main, strings, progress)
            rows = source.findall(f'{NS}sheetData/{NS}row')
            labels = headers(next(r for r in rows if r.get('r') == '6'), strings,
                             {'type','activity id','activity name','wbs','amount'})
            numbers = [int(p.period_id[1:]) for p in payment.periods]
            if numbers != sorted(set(numbers)):
                raise PaymentWorkbookError('Ambiguous or out-of-order period IDs. Review Payment Input.')
            tree, leaves, diagnostics = source_tree(rows, labels, plan, strings, payment.periods)
            resolved = PaymentPositionEngine().resolve(payment, progress)
            periods = []
            for p, position in zip(payment.periods, resolved.periods):
                issues = [f'{s.code}: {s.activity_id}' for s in resolved.issues if s.period_id == p.period_id]
                issues.extend(f'UnverifiedPlan: {r.activity_id}' for r in p.requirements
                              if r.activity_id in unreliable)
                if not p.requirements:
                    issues.append('No explicit requirements to resolve.')
                periods.append(Period(p.period_id, None if issues else position.planned_eligible_date,
                                      p.payment_date, tuple(issues)))
            latest = None
            for i, period in enumerate(periods):
                if period.eligible is None:
                    continue
                if latest is not None and period.eligible < latest:
                    message = f'Computed {period.eligible} precedes {latest}. Review requirement order.'
                    periods[i] = replace(period, eligible=None, issues=(message,))
                else:
                    latest = period.eligible
            return ReportSource(info['Project'], path.name, revision, origin, tuple(periods),
                                tree, leaves, tuple(diagnostics))
    except PaymentWorkbookError:
        raise
    except (BadZipFile, ET.ParseError, ValueError, KeyError, InvalidOperation, StopIteration) as exc:
        raise PaymentWorkbookError(f'Cannot read Payment Plan source. Review and re-prepare: {exc}') from exc
