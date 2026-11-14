"""TR-1 snapshot indexes and view selection; no payment recalculation."""
from decimal import Decimal
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink
from progress_studio.infrastructure.excel.styles import HEADER_FILL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from progress_studio.domain.boq_mapping_trace import mapping_trace
from progress_studio.domain.boq_payment_projection import TOLERANCE, number


def prepare_mapping_trace(workbook, projection, cancelled=lambda: False):
    """Append isolated values-only tables after all accepted BP snapshot blocks."""
    model = mapping_trace(projection)
    data = workbook['_BP_Data']
    row = data.max_row + 3

    def table(name, records, width):
        nonlocal row
        start = row
        records = tuple(records)
        if not records:
            records = (('No saved Mapping',) + ('',) * (width - 1),)
        for record in records:
            if cancelled():
                raise InterruptedError('BOQ export cancelled')
            for column, value in enumerate(record, 1):
                cell = data.cell(row, column, float(value) if isinstance(value, Decimal) else value)
                if isinstance(value, str):
                    cell.data_type = 's'
            row += 1
        from openpyxl.utils import get_column_letter
        workbook.defined_names.add(DefinedName(name, attr_text=
            f"'_BP_Data'!$A${start}:${get_column_letter(width)}${row - 1}"))
        row += 2
        return start, row - 3

    leaf = {entry.item.key: entry for entry in projection.items if entry.item.role == 'item'}
    amounts = []
    records = []
    for edge in model.relationships:
        basis = leaf[edge.boq_key].item.amount
        allocation = number(basis) * edge.share if basis is not None and basis >= 0 else None
        amounts.append(allocation)
        records.append((edge.boq_key, edge.boq_description, edge.activity_id,
                        edge.activity_name, edge.share, 'Unavailable' if allocation is None else allocation))
    table('BP_Relationships', records, 6)

    def total_amount(indexes):
        if any(amounts[index] is None for index in indexes):
            return 'Unavailable'
        return sum((amounts[index] for index in indexes), Decimal(0))

    forward_rows, forward_meta = [], []
    for key, indexes in model.by_boq:
        entry = leaf[key]
        coverage = entry.periods[0].mapping_coverage if entry.periods else Decimal(0)
        if not indexes:
            status = 'Manual BOQ / No Activity Mapping' if getattr(entry, 'origin', '') == 'Manual BOQ' else 'Unmapped'
        else:
            status = 'Full Mapping' if abs(coverage - 1) <= TOLERANCE else f'Partial Mapping ({coverage:.1%})'
        forward_meta.append((len(forward_rows) + 1, len(indexes), status, coverage, total_amount(indexes)))
        forward_rows.extend((index + 1,) for index in indexes)
    table('BP_RelBOQ', forward_meta, 5)
    table('BP_RelBOQRows', forward_rows, 1)
    reverse_rows, activity_meta = [], []
    by_activity = dict(model.by_activity)
    for identity, name in model.activities:
        indexes = by_activity[identity]
        activity_meta.append((identity + ' | ' + name, identity, name, len(reverse_rows) + 1, len(indexes), total_amount(indexes)))
        reverse_rows.extend((index + 1,) for index in indexes)
    if not activity_meta:
        activity_meta = [('No Activities available', '', '', 1, 0, 0)]
    start, end = table('BP_RelActivities', activity_meta, 6)
    workbook.defined_names.add(DefinedName('BP_Activities', attr_text=f"'_BP_Data'!$A${start}:$A${end}"))
    table('BP_RelActivityRows', reverse_rows, 1)
    return model, activity_meta[0][0]


def render_mapping_sheets(workbook, model, first_activity, first_boq):
    """Independent native selectors over the same immutable relationship table."""
    for reverse, name in ((False, 'BOQ Mapping'), (True, 'Activity Mapping')):
        ws = workbook.create_sheet(name, 3 if reverse else 2)
        ws['A1'] = 'Saved Mapping — ' + ('Activity to BOQ Items' if reverse else 'BOQ to Activities')
        ws.merge_cells('A1:C1')
        ws['A2'] = 'Activity ID | Name' if reverse else 'BOQ Item'
        ws['B2'] = first_activity if reverse else first_boq
        ws.merge_cells('B2:C2')
        ws['B2'].alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[2].height = 32
        source = 'BP_Activities' if reverse else 'BP_Items'
        validation = DataValidation(type='list', formula1='=' + source, allow_blank=False)
        validation.showErrorMessage = True
        validation.errorStyle = 'stop'
        validation.error = 'Choose a listed value.'
        ws.add_data_validation(validation)
        validation.add(ws['B2'])
        other = 'BOQ Mapping' if reverse else 'Activity Mapping'
        for address, target in (('A3', other), ('B3', 'Amount'), ('C3', 'Percentage')):
            ws[address] = target
            ws[address].hyperlink = Hyperlink(ref=address, location="'" + target + "'!A1")
            ws[address].style = 'Hyperlink'
        ws['A4'] = 'Select an identity, then F9. Navigation links open a sheet; they do not transfer the selection.'
        ws['A5'] = 'Saved share = % of each BOQ Item allocated to the Activity; never normalized by Activity. No Payment Period required.'
        for row in (4, 5, 6):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical='center')
            ws.row_dimensions[row].height = 30
        table = 'BP_RelActivities' if reverse else 'BP_RelBOQ'
        vector = 'BP_RelActivityRows' if reverse else 'BP_RelBOQRows'
        ws['P2'] = f'=IFERROR(MATCH(B2,{source},0),0)'
        ws['P3'] = f'=IF(P2=0,0,INDEX({table},P2,{5 if reverse else 2}))'
        ws['P4'] = f'=IF(P2=0,1,INDEX({table},P2,{4 if reverse else 1}))'
        status = ('IF(P3=0,"No saved Mapping",P3&" saved BOQ relationships; shares are per BOQ Item")'
                  if reverse else 'INDEX(BP_RelBOQ,P2,3)&" | "&P3&" saved relationships"')
        ws['A6'] = f'=IF(P2=0,"Invalid selector",{status})'
        headers = ('Stable BOQ Item identity', 'BOQ Item description', 'Saved Mapping Share (%)') if reverse else ('Activity ID', 'Activity Name', 'Saved Mapping Share (%)')
        for col, header in enumerate((*headers, 'Allocated Amount'), 1):
            cell = ws.cell(8, col, header)
            cell.fill = PatternFill('solid', fgColor=HEADER_FILL)
            cell.font = Font(color='FFFFFF', bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical='center')
        indexes = model.by_activity if reverse else model.by_boq
        capacity = max(1, max((len(rows) for _, rows in indexes), default=0))
        fields = (1, 2, 5, 6) if reverse else (3, 4, 5, 6)
        money_format = '#,##0.00;[Red]-#,##0.00;0.00'
        if not reverse:
            ws['A7'] = 'Total Mapping Share'
            ws['B7'] = '=IF(P2=0,"Invalid selector",INDEX(BP_RelBOQ,P2,4))'
            ws['B7'].number_format = '0.0%;[Red]-0.0%;0.0%'
        ws['C7'] = 'Total Amount'
        ws['D7'] = f'=IF(P2=0,"Invalid selector",INDEX({table},P2,{6 if reverse else 5}))'
        ws['D7'].number_format = money_format
        for cell in ws[7][:4]:
            cell.font = Font(bold=True)

        from math import ceil
        longest = max((len(e.boq_description if reverse else e.activity_name) for e in model.relationships), default=0)
        height = min(409, max(30, 15 * ceil(longest / 65)))
        for index in range(capacity):
            for col, field in enumerate(fields, 1):
                ws.cell(9 + index, col, f'=IF($P$2=0,"Invalid selector",IF({index}>=$P$3,"",INDEX(BP_Relationships,INDEX({vector},$P$4+{index}),{field})))')
            ws.cell(9 + index, 2).alignment = Alignment(wrap_text=True, vertical='center')
            ws.cell(9 + index, 4).number_format = money_format
            ws.cell(9 + index, 3).number_format = '0.0%;[Red]-0.0%;0.0%'
            ws.row_dimensions[9 + index].height = height
        ws.column_dimensions['A'].width = 32
        ws.column_dimensions['B'].width = 85
        ws.column_dimensions['C'].width = 24
        ws.column_dimensions['D'].width = 24
        ws.column_dimensions['P'].hidden = True
        ws.row_dimensions[8].height = 36
        ws.freeze_panes = 'B9'
        ws.sheet_view.showGridLines = False
        ws.print_area = f'A1:D{8 + capacity}'
        ws.print_title_rows = '1:8'
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
