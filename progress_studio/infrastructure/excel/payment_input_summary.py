"""Locked cumulative WBS display, independent of sparse Payment requirements."""
from math import isfinite

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter

# Only an exact numerical zero is hidden. A small positive number still renders.
WBS_PERCENT_FORMAT = '0.0%;-0.0%;;@'


def project_summary_tree(rows):
    """Retain workbook project metadata, or use an explicitly labelled display root.

    A source WBS enclosing every Activity already provides a genuine root when
    no Project Summary metadata exists. Never fabricate a WBS/Activity ID.
    """
    projects = [dict(r) for r in rows if r['row_type'] == 'PROJECT']
    body = [dict(r) for r in rows if r['row_type'] != 'PROJECT']
    if not projects and body and body[0]['row_type'] == 'WBS' and all(
            r['outline_level'] > body[0]['outline_level'] for r in body[1:]):
        root = body.pop(0)
    else:
        root = projects[0] if projects else {
            'wbs': '', 'activity_name': 'Project Summary', 'weekly_plan': ()}
    root.update(row_type='WBS', activity_id='', outline_level=0)
    root['activity_name'] = root.get('activity_name') or 'Project Summary'
    # Keep existing positive levels. Legacy level-zero branches become children
    # through a uniform shift, preserving all relative hierarchy relationships.
    shift = max(0, 1 - min((r['outline_level'] for r in body), default=1))
    for row in body:
        row['outline_level'] += shift
    return [root] + body


def amount_reference(value):
    """Keep genuine nonnegative numbers; unavailable weights are explicit text."""
    if value is None or value == "":
        return "Missing Amount"
    if isinstance(value, str) and value.startswith("="):
        return "Formula unavailable"
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return "Invalid Amount"
    if not isfinite(value) or value < 0:
        return "Invalid Amount"
    return value


def add_amount_summaries(ws, tree_rows, period_count):
    """Build display helpers without inserting rows or writing Activity inputs.

    Whole-schedule validity is checked before deriving a state. For a known
    schedule, blanks before its first entry mean zero; later blanks retain its
    latest entry. An entirely absent positive-weight schedule stays unknown.
    """
    first, last = 8, 7 + len(tree_rows)
    helper = max(period_count + 6, 9)
    status_col, count_col = helper, helper + 1
    normal_start = helper + 2
    state_start = normal_start + period_count
    contribution_start = state_start + period_count
    decrease_start = contribution_start + period_count
    sl, nl = get_column_letter(status_col), get_column_letter(count_col)
    amber = PatternFill('solid', fgColor='FFF2CC')
    for col in range(helper, decrease_start + period_count):
        ws.column_dimensions[get_column_letter(col)].hidden = True
    ws.cell(6, status_col, 'Activity diagnostic')
    ws.cell(6, count_col, 'Entered requirements')
    for offset in range(period_count):
        period = ws.cell(6, offset + 6).value
        for start, name in ((normal_start, 'normalized input'), (state_start, 'derived state'),
                            (contribution_start, 'weighted state'), (decrease_start, 'decrease warning')):
            ws.cell(6, start + offset, f'{period} {name}')

    for row, item in enumerate(tree_rows, first):
        level = max(int(item.get('outline_level') or 0), 0)
        ws.row_dimensions[row].outlineLevel = min(level, 7)
        ws.row_dimensions[row].height = 22
        if item['row_type'] != 'ACT':
            continue
        ws.cell(row, 5, item.get('amount', 'Missing Amount')).number_format = '#,##0.00'
        normals = f'{get_column_letter(normal_start)}{row}:{get_column_letter(state_start-1)}{row}'
        ws.cell(row, count_col, f'=COUNT({normals})')
        ws.cell(row, status_col,
                f'=IF(NOT(ISNUMBER(E{row})),"Check Amount",IF(COUNTIF({normals},"Check input")>0,'
                f'"Check input",IF(AND({nl}{row}=0,E{row}>0),"Incomplete","")))')
        for offset in range(period_count):
            raw = f'{get_column_letter(6+offset)}{row}'
            normal = f'{get_column_letter(normal_start+offset)}{row}'
            state = f'{get_column_letter(state_start+offset)}{row}'
            previous = f'{get_column_letter(state_start+offset-1)}{row}' if offset else '0'
            previous_warning = f'{get_column_letter(decrease_start+offset-1)}{row}' if offset else '0'
            # The sparse reader trims whitespace and accepts numeric text / %.
            ws.cell(row, normal_start + offset,
                    f'=IFERROR(IF(LEN(TRIM({raw}&""))=0,"",IF(AND({raw}*1>=0,{raw}*1<=1),'
                    f'{raw}*1,"Check input")),"Check input")')
            ws.cell(row, state_start + offset,
                    f'=IF(COUNTIF({normals},"Check input")>0,"Check input",'
                    f'IF({nl}{row}=0,"",IF(ISNUMBER({normal}),{normal},{previous})))')
            ws.cell(row, contribution_start + offset,
                    f'=IFERROR(IF(AND(ISNUMBER({state}),ISNUMBER(E{row})),{state}*E{row},0),0)')
            ws.cell(row, decrease_start + offset,
                    f'=IFERROR(IF(AND(ISNUMBER({normal}),ISNUMBER({previous}),'
                    f'{normal}<{previous}),1,{previous_warning}),{previous_warning})')

    # One relative rule covers all Activity inputs, rather than one rule per
    # cell (important for large workbooks and the supported 120 periods).
    first_normal = f'{get_column_letter(normal_start)}{first}'
    prior = f'IF(COLUMN()=6,0,{get_column_letter(state_start-1)}{first})'
    if tree_rows:
        ws.conditional_formatting.add(f'F{first}:{get_column_letter(period_count+5)}{last}', FormulaRule(
            formula=[f'IFERROR(AND($A{first}="ACT",ISNUMBER({first_normal}),ISNUMBER({prior}),{first_normal}<{prior}),FALSE)'],
            fill=amber))

    for index, item in enumerate(tree_rows):
        if item['row_type'] != 'WBS':
            continue
        row = first + index
        end_index = index + 1
        while end_index < len(tree_rows) and tree_rows[end_index]['outline_level'] > item['outline_level']:
            end_index += 1
        descendants = [t for t in tree_rows[index+1:end_index] if t['row_type'] == 'ACT']
        weights = [t.get('amount', 'Missing Amount') for t in descendants]
        valid = all(isinstance(w, (int, float)) and not isinstance(w, bool) and isfinite(w) and w >= 0 for w in weights)
        total = sum(weights) if valid else 'Check Amount'
        if valid and not isfinite(total):
            total = 'Check Amount'
        ws.cell(row, 5, total).number_format = '#,##0.00'
        start, end = row + 1, first + end_index - 1
        for offset in range(period_count):
            cell = ws.cell(row, 6 + offset)
            cell.number_format = WBS_PERCENT_FORMAT
            if not descendants:
                continue
            statuses = f'${sl}${start}:${sl}${end}'
            counts = f'${nl}${start}:${nl}${end}'
            contribution = get_column_letter(contribution_start + offset)
            warning = get_column_letter(decrease_start + offset)
            cell.value = (
                f'=IF(NOT(ISNUMBER($E{row})),"Check Amount",'
                f'IF(COUNTIF({statuses},"Check input")>0,"Check input",'
                f'IF(SUM({counts})=0,"",IF($E{row}=0,"No weight",'
                f'IF(COUNTIF({statuses},"Incomplete")>0,"Incomplete",'
                f'SUM({contribution}${start}:{contribution}${end})/$E{row})))))')
        if descendants:
            warning = get_column_letter(decrease_start)
            if period_count > 1:
                # Presentation only: compare adjacent actual numbers, never
                # rounded percentages. Preserve the existing amber warning
                # even when its numerical value is visually suppressed.
                complete = f'AND(ISNUMBER(G{row}),G{row}=1,ISNUMBER(F{row}),F{row}=1)'
                next_warning = get_column_letter(decrease_start + 1)
                # Include the existing warning fill in the first matching
                # completion rule for readers that do not combine CF styles.
                for condition, fill in (
                    (f'AND({complete},SUM({next_warning}${start}:{next_warning}${end})>0)', amber),
                    (complete, None),
                ):
                    rule = FormulaRule(formula=[f'IFERROR({condition},FALSE)'], stopIfTrue=False)
                    rule.dxf = DifferentialStyle(
                        numFmt=NumberFormat(numFmtId=164, formatCode=';;;'), fill=fill)
                    ws.conditional_formatting.add(
                        f'G{row}:{get_column_letter(period_count+5)}{row}', rule)
            ws.conditional_formatting.add(f'F{row}:{get_column_letter(period_count+5)}{row}', FormulaRule(
                formula=[f'SUM({warning}${start}:{warning}${end})>0'], fill=amber, stopIfTrue=False))

    # A single fixed diagnostic line replaces repetitive coverage rows. Include
    # the first affected Activity ID so the user can find it without unhiding data.
    statuses = f'${sl}${first}:${sl}${max(first,last)}'
    ids = f'$C${first}:$C${max(first,last)}'
    final_warning = get_column_letter(decrease_start + period_count - 1)
    warnings = f'${final_warning}${first}:${final_warning}${max(first,last)}'
    diagnostic = '"OK"'
    diagnostic = f'IF(SUM({warnings})>0,"Review decreasing requirements: "&INDEX({ids},MATCH(1,{warnings},0))&" (amber cells)",{diagnostic})'
    diagnostic = f'IF(SUM({nl}{first}:{nl}{max(first,last)})=0,"No requirements",{diagnostic})'
    for label in ('Incomplete', 'Check input', 'Check Amount'):
        diagnostic = f'IF(COUNTIF({statuses},"{label}")>0,"{label}: "&INDEX({ids},MATCH("{label}",{statuses},0)),{diagnostic})'
        if label == 'Incomplete':
            diagnostic = f'IF(SUM({nl}{first}:{nl}{max(first,last)})=0,"No requirements",{diagnostic})'
    ws['D4'] = '=' + diagnostic
    ws.freeze_panes = 'F8'
    for span in ('A5:C5', 'D5:H5', 'D2:H2', 'D3:H3', 'D4:H4'):
        ws.merge_cells(span)
    ws['A5'].alignment = Alignment(wrap_text=True, vertical='center')
    ws['D5'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[5].height = 30
    ws.column_dimensions['E'].width = 23
    for col in range(6, period_count + 6):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws['D2'] = 'Cumulative WBS Requirement (Amount-weighted). F9 / Save; exact zero and consecutive exact 100% are hidden.'
    ws['D3'] = 'WBS only: before first entry = 0; later blanks retain latest; never entered = unknown. Amount refresh: Prepare.'
    for row in (2, 3, 4):
        ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical='center')
        ws.cell(row, 4).alignment = Alignment(wrap_text=True, vertical='center')
        ws.cell(row, 4).font = Font(size=10, color='595959')
        ws.row_dimensions[row].height = 30
    ws.print_options.horizontalCentered = True
    ws.print_area = f'A1:{get_column_letter(period_count+5)}{max(first,last)}'
    ws.print_title_rows = '1:7'
