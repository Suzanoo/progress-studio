"""Versioned, persistent BOQ input. Formula cells are never accepted as inputs."""
from decimal import Decimal
import hashlib
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Protection, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from progress_studio.domain.boq_manual_requirements import ManualInput, manual_periods
from progress_studio.infrastructure.excel.styles import HEADER_FILL
from progress_studio.config.workbook_protection import WORKBOOK_SHEET_PASSWORD

SHEET = 'BOQ Requirement Input'
SCHEMA = 'bp4-signed-period-v1'
FIRST_ROW, FIRST_COLUMN = 9, 9


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode('utf-8')).hexdigest()


def signature(item):
    return digest((item.key, item.stable_id, item.source_sheet, item.source_row,
                   item.wbs, item.description, item.hierarchy))


def metadata(projection):
    return {'schema': SCHEMA, 'catalog': digest([(e.item.key, signature(e.item), e.item.amount)
                for e in projection.items]),
            'allocations': digest([(e.item.key, [(c.activity_id, c.share) for c in e.periods[0].contributions])
                for e in projection.items if e.periods]),
            'periods': [(s.period_id, str(s.payment_date)) for s in projection.period_sources],
            'lineage': dict(projection.provenance)}


def eligible(projection):
    return [e for e in projection.items if e.item.role == 'item'
            and all(not p.contributions and p.mapping_coverage == 0 for p in e.periods)]


def read_input(path, projection, *, reconcile=False):
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if SHEET not in wb.sheetnames:
            return None
        ws = wb[SHEET]
        try:
            saved = json.loads(ws['B2'].value)
        except (TypeError, ValueError) as exc:
            raise ValueError('Invalid BOQ Requirement Input metadata') from exc
        current = metadata(projection)
        if saved.get('schema') != SCHEMA:
            raise ValueError('Unsupported BOQ Requirement Input schema')
        if saved.get('periods') != json.loads(json.dumps(current['periods'])):
            raise ValueError('BOQ input period identity/date conflict; preserve source and resolve explicitly')
        if not reconcile and any(saved.get(k) != current[k] for k in ('catalog', 'allocations')):
            raise ValueError('BOQ input source revision changed; run Prepare/Reconcile first')
        headers = tuple(ws.cell(7, FIRST_COLUMN + i).value for i in range(len(projection.periods)))
        if headers != projection.periods:
            raise ValueError('BOQ input period headers changed')
        entries = {e.item.key: e for e in eligible(projection)}
        records, seen = [], set()
        for cells in ws.iter_rows(min_row=FIRST_ROW):
            if all(c.value is None for c in cells):
                continue
            key = cells[0].value
            if key in seen or key not in entries:
                raise ValueError('BOQ input identity missing, duplicate or now mapped: ' + str(key))
            seen.add(key);entry = entries[key]
            if cells[1].value != entry.item.stable_id or cells[2].value != signature(entry.item):
                raise ValueError('BOQ input identity signature changed: ' + key)
            if any(c.value is not None for c in cells[FIRST_COLUMN - 1 + len(projection.periods):]):
                raise ValueError('Unexpected BOQ input period data')
            values = []
            for c in cells[FIRST_COLUMN - 1:FIRST_COLUMN - 1 + len(projection.periods)]:
                value = c.value
                if value is not None and (c.data_type == 'f' or isinstance(value, bool) or not isinstance(value, (float, int, Decimal))):
                    raise ValueError(f'Invalid manual numeric input at {SHEET}!{c.coordinate}')
                values.append(None if value is None else Decimal(str(value)))
            record = ManualInput(key, tuple(values), cells[0].row)
            manual_periods(entry, record, projection.periods)
            records.append(record)
        if not reconcile and seen != set(entries):
            raise ValueError('BOQ input population changed; run Prepare/Reconcile first')
        return tuple(records)
    finally:
        wb.close()


def render_input(projection, records):
    wb = Workbook();ws = wb.active;ws.title = SHEET
    ws['D1'] = 'BOQ Planned Payment Requirement — signed Period %, not certified payment'
    ws['A2'] = 'Schema / source revision';ws['B2'] = json.dumps(metadata(projection), ensure_ascii=False, default=str)
    ws.row_dimensions[2].hidden = True
    ws['D3'] = 'Enter signed Period % of FULL BOQ Amount. Blank = unknown; 0% = confirmed zero.'
    ws['D4'] = 'Save this source workbook, then Export BOQ Payment Workbook. Activity Payment Rebuild is not required.'
    ws['D5'] = 'Only wholly unmapped items. Do not change identities or period headers. Prepare preserves inputs or stops on conflicts.'
    headers = ('BOQ key','Stable ID','Identity signature','WBS','BOQ Item','Unit','BOQ Amount','Origin / status') + projection.periods
    for c,v in enumerate(headers,1):
        cell=ws.cell(7,c,v);cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor=HEADER_FILL)
    for i,p in enumerate(projection.period_sources):
        ws.cell(8,FIRST_COLUMN+i,p.payment_date).number_format='dd-mmm-yy'
    prior = {r.key: r.values for r in records or ()}
    for r,e in enumerate(eligible(projection),FIRST_ROW):
        item=e.item
        for c,v in enumerate((item.key,item.stable_id,signature(item),item.wbs,item.description,item.unit,
                              float(item.amount) if item.amount is not None else 'Unavailable',
                              'Manual BOQ / Unmapped'),1):
            cell=ws.cell(r,c,v)
            if isinstance(v,str):cell.data_type='s'
        ws.cell(r,7).number_format='#,##0.00'
        for i,v in enumerate(prior.get(item.key,(None,)*len(projection.periods))):
            cell=ws.cell(r,FIRST_COLUMN+i,None if v is None else float(v))
            cell.number_format='0.0%;[Red]-0.0%;0.0%';cell.protection=Protection(locked=False)
    last = max(FIRST_ROW,ws.max_row)
    from openpyxl.utils import get_column_letter
    dv=DataValidation(type='decimal',operator='between',formula1='-1',formula2='1',allow_blank=True)
    dv.showErrorMessage=True;dv.errorStyle='stop';dv.error='Enter a numeric signed Period % from -100% to 100%, or leave blank.'
    ws.add_data_validation(dv);dv.add(f'I{FIRST_ROW}:{get_column_letter(FIRST_COLUMN+len(projection.periods)-1)}{last}')
    for c in ('A','B','C'):ws.column_dimensions[c].hidden=True
    for c,width in (('D',18),('E',55),('F',12),('G',22),('H',30)):ws.column_dimensions[c].width=width
    for c in range(FIRST_COLUMN,FIRST_COLUMN+len(projection.periods)):ws.column_dimensions[get_column_letter(c)].width=16
    ws.freeze_panes='I9';ws.sheet_view.showGridLines=False
    ws.protection.sheet=True;ws.protection.password=WORKBOOK_SHEET_PASSWORD
    ws.print_title_rows='1:8';ws.print_title_cols='D:H';ws.page_setup.orientation='landscape'
    return wb
