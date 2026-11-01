"""WU-2 calendar-day display. Never supplies weighting, BAC or cutoff inputs."""
from datetime import date, datetime
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils.datetime import to_excel, WINDOWS_EPOCH
from progress_studio.infrastructure.excel.amount_workbook import find_header

MILESTONE_PREFIX = 'Calendar Duration Milestone: '
NOTE = ('Inclusive calendar days, not working duration or weight. Date times are ignored. '
        'Blank means missing/invalid/reversed dates or incomplete Actual. Explicit milestones = 0. '
        'Plan summaries use their existing date span; Actual summaries stay blank. '
        'Edit dates in main then F9/Save; Progress Rebuild refreshes copied Monthly metadata.')


def milestone_ids(workbook):
    if 'Info' not in workbook.sheetnames:return set()
    return {str(row[0].value)[len(MILESTONE_PREFIX):].upper()
            for row in workbook['Info'].iter_rows(min_col=1,max_col=2)
            if str(row[0].value or '').startswith(MILESTONE_PREFIX) and row[1].value is True}


def seed_milestones(workbook, activities):
    for activity in activities:
        if not activity.is_summary and activity.is_milestone:
            workbook['Info'].append([MILESTONE_PREFIX+activity.activity_id.strip().upper(),True])


def layout(ws):
    row,h=find_header(ws,['Row Type','Description'])
    present=[x in h for x in ('plan duration','actual duration')]
    if not any(present):return None
    if not all(present) or not all(x in h for x in ('p/a','plan start','plan finish','actual start','actual finish')) or h['plan duration']+1!=h['plan start'] or h['actual duration']+1!=h.get('actual start'):
        raise ValueError('Unsupported Duration column layout; no automatic migration is performed.')
    return row,h


def calendar_days(start, finish, *, milestone=False, epoch=WINDOWS_EPOCH):
    def serial(value):
        if isinstance(value,datetime):return to_excel(value,epoch)
        if isinstance(value,date):return to_excel(datetime.combine(value,datetime.min.time()),epoch)
        if isinstance(value,(int,float)) and not isinstance(value,bool):return value
        return None
    s,f=serial(start),serial(finish)
    if s is None or f is None or not (1<=s<2958466 and 1<=f<2958466) or int(f)<int(s):return None
    return 0 if milestone else int(f)-int(s)+1


def populate_durations(workbook, target=None, *, snapshot=False, value_source=None):
    """Only existing WU-2 columns are populated; legacy sheets are untouched."""
    source=workbook['main'];resolved=layout(source)
    if resolved is None:return
    header,h=resolved;target=source if target is None else target
    milestones=milestone_ids(workbook)
    existing_warnings={tuple(rule.formula or ()) for area in target.conditional_formatting for rule in target.conditional_formatting[area]}
    for name in ('plan duration','actual duration'):
        target.cell(header,h[name]).comment=Comment(NOTE,'Progress Studio')
    for row in range(header+1,source.max_row+1):
        kind=str(source.cell(row,h['row type']).value or '').strip().lower()
        pa=str(source.cell(row,h['p/a']).value or '').upper()
        if kind not in ('activity','wbs','project summary') or pa!='P':continue
        aid=str(source.cell(row,h.get('activity id',1)).value or '').strip().upper()
        milestone=kind=='activity' and aid in milestones
        for name,start,finish,r,enabled in (
            ('plan duration','plan start','plan finish',row,True),
            ('actual duration','actual start','actual finish',row+1,kind=='activity'),
        ):
            col=h[name];cell=target.cell(r,col)
            target.cell(row+1 if name=='plan duration' else row,col).value=None
            cell.number_format='0" d"';cell.protection=Protection(locked=True)
            cell.alignment=Alignment(horizontal='center')
            if not enabled:cell.value=None;continue
            s=source.cell(r,h[start]);f=source.cell(r,h[finish])
            if snapshot:
                values=[value_source.cell(r,c.column).value if c.data_type=='f' and value_source is not None else c.value for c in (s,f)]
                cell.value=calendar_days(*values,milestone=milestone,epoch=workbook.epoch)
            else:
                a,b=s.coordinate,f.coordinate
                valid=f'AND(ISNUMBER({a}),ISNUMBER({b}),{a}>=1,{b}>=1,{a}<2958466,{b}<2958466,INT({b})>=INT({a}))'
                value='0' if milestone else f'INT({b})-INT({a})+1'
                cell.value=f'=IFERROR(IF({valid},{value},""),"")'
                # A visible indication when both dates are entered but unusable.
                if target is source:
                    warning=f'IFERROR(AND({a}<>"",{b}<>"",NOT({valid})),TRUE)'
                    # Renderer owns only this exact single-cell warning/formula.
                    rules=target.conditional_formatting
                    if (warning,) not in existing_warnings:
                        rules.add(cell.coordinate,FormulaRule(formula=[warning],fill=PatternFill('solid',fgColor='FFC7CE')))
                        existing_warnings.add((warning,))
