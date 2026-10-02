"""Workbook boundary for independent Contract Value and milestone history.

The input sheet is persistent user data, never an EV-generated snapshot.
"""
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import math
from openpyxl.styles import Font, PatternFill, Protection

SHEET = 'EV Monetary Inputs'
ACTIVITY = 'Activity Amount'
BOQ = 'BOQ Mapping'
HEADERS = ('Activity ID', 'Description', 'Allocated Contract Value', 'Milestone',
           'Planned Milestone Date', 'Actual Completion Date')


def monetary(value, label):
    if value is None or value == '' or isinstance(value, bool):
        raise ValueError(f'{label}: missing or invalid monetary Amount.')
    try:
        amount = Decimal(str(value))
        converted = float(amount)
    except (InvalidOperation, ValueError, TypeError, OverflowError):
        raise ValueError(f'{label}: nonnumeric monetary Amount.') from None
    if not amount.is_finite() or amount < 0 or not math.isfinite(converted):
        raise ValueError(f'{label}: negative or non-finite monetary Amount.')
    if amount != 0 and (abs(converted) < 2.2250738585072014e-308 or abs(converted) > 9.99999999999999e307):
        raise ValueError(f'{label}: monetary Amount outside Excel numeric range.')
    return converted


def _date(value, label, required=False):
    if value in (None, '') and not required:
        return None
    if isinstance(value, datetime):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    raise ValueError(f'{label}: enter an Excel date.')


def write_inputs(workbook, records):
    if SHEET in workbook.sheetnames:
        del workbook[SHEET]
    ws = workbook.create_sheet(SHEET)
    ws.append(['EV monetary basis: Allocated Contract Value. BAC edits require EV refresh.'])
    ws.append(['Milestone dates are live. Blank Actual Completion Date means not completed.'])
    ws.append(HEADERS)
    for record in records:
        ws.append(record)
        for cell in ws[ws.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'
        for col in (3, 4, 5, 6):
            ws.cell(ws.max_row, col).protection = Protection(locked=False)
        ws.cell(ws.max_row, 3).number_format = '#,##0.00'
        for col in (5, 6):
            ws.cell(ws.max_row, col).number_format = 'dd-mmm-yyyy'
    for cell in ws[3]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='1F4E78')
    for col, width in zip('ABCDEF', (20, 48, 28, 14, 26, 26)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'C4'
    ws.auto_filter.ref = f'A3:F{max(4,ws.max_row)}'
    return ws


def seed_import(workbook, rows, weight_basis, field):
    records = []
    for item in rows:
        if item.is_summary:
            continue
        value = None
        if weight_basis == 'amount':
            if item.is_milestone:
                raw = [v for k,v in item.amount_field_values if field and k == field.identity]
                # Preserve an explicit source value, including invalid text, for
                # readiness validation. Never silently substitute zero.
                if len(raw) == 1:
                    try:
                        value = monetary(raw[0], item.activity_id)
                    except ValueError:
                        value = raw[0]
            else:
                value = item.amount
        records.append((item.activity_id, item.name, value, bool(item.is_milestone),
                        item.plan_finish if item.is_milestone else None,
                        item.actual_finish if item.is_milestone else None))
    write_inputs(workbook, records)


def read_inputs(workbook, dataset, *, require_amount):
    if SHEET not in workbook.sheetnames:
        if require_amount:
            raise ValueError(f"Missing '{SHEET}'. Prepare explicit Contract Values before Activity Amount EV; dummy weights are not BAC.")
        return {}, ()
    ws = workbook[SHEET]
    if tuple(ws.cell(3,c).value for c in range(1,7)) != HEADERS:
        raise ValueError(f'{SHEET}: invalid input headers.')
    rows = {}
    for row in range(4,ws.max_row+1):
        identity = str(ws.cell(row,1).value or '').strip()
        if not identity:
            continue
        if identity in rows:
            raise ValueError(f'Duplicate Activity ID in {SHEET}: {identity}')
        rows[identity] = row
    amounts, milestones = {}, []
    for activity in dataset.activities:
        identity = activity.activity_id.strip()
        if identity not in rows:
            raise ValueError(f'{SHEET}: missing Activity ID {identity}.')
        row = rows[identity]
        if require_amount:
            amounts[identity] = monetary(ws.cell(row,3).value, identity)
        flag = ws.cell(row,4).value
        if flag not in (True, False, 0, 1):
            raise ValueError(f'{identity}: Milestone must be TRUE or FALSE.')
        if flag:
            if activity.amount != 0:
                raise ValueError(f'{identity}: milestone Progress Weight must remain zero.')
            planned = _date(ws.cell(row,5).value, f'{identity} Planned Milestone Date', True)
            actual = _date(ws.cell(row,6).value, f'{identity} Actual Completion Date')
            milestones.append((identity, planned, actual, row))
    return amounts, tuple(milestones)


def prepare_inputs(source, output):
    """Explicit preparation for legacy workbooks; never infer monetary BAC.

    The user supplies Contract Values and milestone identity/dates in Excel.
    New Create workbooks already carry these fields from their chosen XML input.
    """
    from pathlib import Path
    from openpyxl import load_workbook
    from progress_studio.infrastructure.excel.rebuild_workbook_reader import RebuildWorkbookReader
    if Path(source).resolve() == Path(output).resolve():
        raise ValueError('Select a new output workbook.')
    dataset = RebuildWorkbookReader().read_main_dataset(Path(source))
    wb = load_workbook(source, keep_vba=Path(source).suffix.lower()=='.xlsm')
    try:
        if SHEET in wb.sheetnames:
            raise ValueError(f'{SHEET} already exists. Edit it in Excel, then refresh EV.')
        from progress_studio.infrastructure.excel.weight_basis import creation_basis
        # Provenance is only an eligibility gate. Current main Amount is read,
        # never the historic XML value/metadata payload. Dummy bases stay blank.
        is_monetary = creation_basis(wb) == 'amount'
        write_inputs(wb, [(a.activity_id, a.description, a.amount if is_monetary else None,
                           False, None, None) for a in dataset.activities])
        wb.save(output)
    finally:
        wb.close()


def resolve_inputs(path, dataset, monetary_source, input_reader):
    from dataclasses import replace
    from math import isclose, isfinite
    from openpyxl import load_workbook
    if any(row.row_type.lower() == 'activity' and row.pa.upper() == 'P'
           and not row.activity_id.strip() for row in dataset.rows):
        raise ValueError('Missing Activity ID in main Plan row.')
    ids = [a.activity_id.strip() for a in dataset.activities]
    if len(ids) != len(set(ids)):
        raise ValueError('Duplicate Activity ID in main Plan rows.')
    for pa in ('P', 'A'):
        seen = set()
        for row in dataset.rows:
            if row.pa.upper() == pa and row.activity_id:
                identity = row.activity_id.strip()
                if identity in seen:
                    raise ValueError(f'Duplicate Activity ID {identity} ({pa}).')
                seen.add(identity)
    workbook = load_workbook(path, data_only=False)
    try:
        amounts, milestones = read_inputs(workbook, dataset, require_amount=monetary_source == ACTIVITY)
        if monetary_source == ACTIVITY:
            from progress_studio.infrastructure.excel.earned_value_input_reader import EmbeddedEarnedValueInputs
            embedded = EmbeddedEarnedValueInputs(path, (), ())
        else:
            embedded = input_reader.read(path)
            totals = {identity: 0.0 for identity in ids}
            boq_by_key = {b.key: b for b in embedded.boq_rows}
            for allocation in embedded.allocations:
                value = monetary(boq_by_key[allocation.boq_key].amount, allocation.boq_key)
                share = monetary(allocation.share_percent, allocation.boq_key + ' share')
                if share > 100:
                    raise ValueError('BOQ share exceeds 100%.')
                key = allocation.activity_id.strip()
                if key not in totals:
                    raise ValueError(f'Unknown Activity ID {key}.')
                totals[key] += value * share / 100
            milestone_ids = {m[0] for m in milestones}
            for activity in dataset.activities:
                key = activity.activity_id.strip()
                # Milestone weight is deliberately zero; BAC comes from
                # reconciled allocation rows, never progress weight.
                if key not in milestone_ids:
                    current = monetary(activity.amount, key)
                    if not isclose(current, totals[key], rel_tol=1e-10, abs_tol=1e-7):
                        raise ValueError(f'BOQ BAC reconciliation mismatch for {key}: main {current}, allocations {totals[key]}.')
            ws = workbook['BOQ Activity Mapping']
            headers = {str(c.value).strip().lower(): c.column for c in ws[1] if c.value}
            if 'allocated amount' not in headers:
                raise ValueError('BOQ reconciliation requires Allocated Amount.')
            for row in range(2, ws.max_row + 1):
                key = str(ws.cell(row, headers['boq key']).value or '').strip()
                if not key:
                    continue
                allocated = monetary(ws.cell(row,headers['allocated amount']).value, key)
                amount = monetary(ws.cell(row,headers['boq amount']).value, key)
                share = monetary(ws.cell(row,headers['share %']).value, key)
                if not isclose(allocated, amount * share, rel_tol=1e-10, abs_tol=1e-7):
                    raise ValueError(f'BOQ Allocated Amount reconciliation mismatch: {key}, row {row}.')
            amounts = totals
    finally:
        workbook.close()
    total = sum(amounts.values())
    if not isfinite(total) or total <= 0 or total > 9.99999999999999e307:
        raise ValueError('EV monetary project BAC must be positive and within Excel numeric range.')
    dataset = replace(dataset, rows=tuple(
        replace(row, amount=amounts[row.activity_id.strip()])
        if row in dataset.activities else row for row in dataset.rows))
    return dataset, embedded, milestones

