"""Create weight provenance in existing Info, with presentation-only labels.

Info is already an internal-preserve sheet in the Rebuild contract. No new
worksheet or cell binding is introduced into the domain/services.
"""
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font

KEY = "Creation Weight Basis"


def set_creation_basis(workbook, basis: str) -> None:
    ws = workbook['Info'] if 'Info' in workbook.sheetnames else workbook.create_sheet('Info')
    for row in ws.iter_rows(min_col=1, max_col=2):
        if row[0].value == KEY:
            row[1].value = basis
            return
    ws.append([KEY, basis])


def creation_basis(workbook) -> str | None:
    if 'Info' not in workbook.sheetnames:
        return None
    for row in workbook['Info'].iter_rows(min_col=1, max_col=2):
        if row[0].value == KEY and row[1].value in {'equal', 'duration'}:
            return row[1].value
    return None


def uses_dummy_weights(workbook) -> bool:
    # Mapping supplies its own real amount source. Keep creation provenance,
    # but never label a subsequently BOQ-mapped workbook as current dummy data.
    return creation_basis(workbook) is not None and 'BOQ Activity Mapping' not in workbook.sheetnames


def apply_weight_labels(workbook) -> None:
    basis = creation_basis(workbook)
    if basis is None:
        return  # legacy files have no new semantics inferred from their values
    dummy = uses_dummy_weights(workbook)
    message = (f"Weight basis: {basis.title()} — dummy units, not Contract Value or financial data."
               if dummy else f"Created with {basis.title()} weights. Current amounts: BOQ Mapping; see Mapping Summary.")
    if 'README' in workbook.sheetnames:
        ws = workbook['README']
        ws.merge_cells('B4:F4')
        ws['B4'] = message
        ws['B4'].font = Font(size=11, bold=True, color='17365D')
        ws['B4'].alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[4].height = 30
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
