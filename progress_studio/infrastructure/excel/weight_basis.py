"""Create weight provenance in existing Info, with presentation-only labels.

Info is already an internal-preserve sheet in the Rebuild contract. No new
worksheet or cell binding is introduced into the domain/services.
"""
from dataclasses import asdict
from progress_studio.domain.amount_field import AmountField
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font

KEY = "Creation Weight Basis"


def _set_info(workbook, key, value):
    ws = workbook['Info'] if 'Info' in workbook.sheetnames else workbook.create_sheet('Info')
    for row in ws.iter_rows(min_col=1, max_col=2):
        if row[0].value == key:
            row[1].value = value
            row[1].data_type = 's'
            return
    ws.append([key, value])
    ws.cell(ws.max_row, 2).data_type = 's'


def _get_info(workbook, key):
    if 'Info' in workbook.sheetnames:
        for row in workbook['Info'].iter_rows(min_col=1, max_col=2):
            if row[0].value == key:
                return row[1].value
    return None


def set_creation_basis(workbook, basis: str) -> None:
    _set_info(workbook, KEY, basis)


def creation_basis(workbook) -> str | None:
    value = _get_info(workbook, KEY)
    return value if value in {'equal', 'duration', 'amount'} else None


def set_creation_field(workbook, field: AmountField) -> None:
    for name, value in asdict(field).items():
        _set_info(workbook, 'Creation Amount Field ' + name, value)


def creation_field(workbook) -> AmountField | None:
    values = {name: _get_info(workbook, 'Creation Amount Field ' + name)
              for name in ('identity', 'name', 'source', 'data_type', 'native_name')}
    if not values['identity']:
        return None
    return AmountField(**{name: value or '' for name, value in values.items()})


def uses_dummy_weights(workbook) -> bool:
    # Mapping supplies its own real amount source. Keep creation provenance,
    # but never label a subsequently BOQ-mapped workbook as current dummy data.
    return creation_basis(workbook) in {'equal', 'duration'} and 'BOQ Activity Mapping' not in workbook.sheetnames


def apply_weight_labels(workbook) -> None:
    basis = creation_basis(workbook)
    if basis is None:
        return  # legacy files have no new semantics inferred from their values
    dummy = uses_dummy_weights(workbook)
    field = creation_field(workbook)
    if 'BOQ Activity Mapping' in workbook.sheetnames:
        message = f"Created with {basis.title()} weights. Current amounts: BOQ Mapping; see Mapping Summary."
    elif dummy:
        message = f"Weight basis: {basis.title()} — dummy units, not Contract Value or financial data."
    else:
        selected = field.label if field else 'see Info metadata'
        message = f"Weight basis: Amount — monetary XML values from {selected}. Not actual cost or cash flow."
    if 'README' in workbook.sheetnames:
        ws = workbook['README']
        ws.merge_cells('B4:F4')
        ws['B4'] = message
        ws['B4'].font = Font(size=11, bold=True, color='17365D')
        ws['B4'].alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[4].height = 54 if basis == "amount" else 30
    # Preserve engine-recognized Amount headers; explain their units in comments.
    for name in ('main', 'main_monthly', 'Amount Mapping', 'Dashboard', 'Payment', 'Payment-Breakdown', 'Earned Value', 'EV Table'):
        if name not in workbook.sheetnames:
            continue
        ws = workbook[name]
        for row in ws.iter_rows(max_row=min(ws.max_row, 40)):
            for cell in row:
                if cell.value in ('Amount', 'XML Amount', 'Total', 'Project Value', 'Total weight'):
                    if cell.comment is None or cell.comment.author == 'Progress Studio Weight':
                        cell.comment = Comment(message, 'Progress Studio Weight')
        if name == 'Dashboard' and ws['J6'].value in ('Project Value', 'Total weight'):
            ws['J6'] = 'Total weight' if dummy else 'Project Value'
        if name == 'Payment-Breakdown' and str(ws['A1'].value or '').startswith('Payment Breakdown'):
            ws['A1'] = (f'Payment Breakdown — {basis.title()} dummy-weighted progress'
                        if dummy else 'Payment Breakdown')
