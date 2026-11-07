"""BP-2 renderer: values-only source data and bounded presentation formulas."""
from decimal import Decimal
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter as letter
from progress_studio.domain.boq_payment_report import report_rows, known_projected_totals
from progress_studio.infrastructure.excel.export_theme import DEFAULT_ACTIVITY_DATA_PALETTE as PALETTE
from progress_studio.infrastructure.excel.styles import HEADER_FILL, PROJECT_FILL, ACTIVITY_PLAN_FILL
from progress_studio.infrastructure.excel.calculation_policy import request_initial_manual_excel_recalculation
from progress_studio.config.workbook_protection import WORKBOOK_SHEET_PASSWORD


def literal(cell, value):
    if isinstance(value, Decimal):value=float(value)
    cell.value=value
    if isinstance(value,str):cell.data_type='s'


def dropdown(ws, cell, formula):
    d=DataValidation(type='list',formula1=formula,allow_blank=False)
    d.showErrorMessage=True;d.errorStyle='stop';d.error='Choose a listed value.'
    ws.add_data_validation(d);d.add(ws[cell])


def render(projection, cancelled=lambda:False):
    rows=report_rows(projection);n=len(rows);periods=len(projection.periods)
    wb=Workbook();wb.remove(wb.active)
    pct=wb.create_sheet('Percentage');amt=wb.create_sheet('Amount');trace=wb.create_sheet('Trace');data=wb.create_sheet('_BP_Data')
    data.sheet_state='hidden';data.protection.sheet=True;data.protection.password=WORKBOOK_SHEET_PASSWORD
    wb.properties.title='Progress Studio BOQ Payment Projection';wb.properties.subject='BP2_DERIVED_REPORT_V1_NOT_BOQ_SOURCE'
    def check():
        if cancelled():raise InterruptedError('BOQ export cancelled')
    # Four vertical blocks. Unavailable cells carry a text sentinel, not empty cells
    # that Excel INDEX could turn into numeric zero.
    for mode in range(4):
        for i,row in enumerate(rows):
            check()
            for p,value in enumerate(row.values[mode]):literal(data.cell(1+mode*n+i,p+1),'Unavailable' if value is None else value)
    leaves=[x for x in projection.items if x.item.role=='item']
    meta=4*n+2;trace_start=meta+len(leaves)+2;cursor=trace_start
    labels=[];maximum=0
    for i,item in enumerate(leaves):
        check();label=f'{item.item.wbs} | {item.item.description} | {item.item.key}';labels.append(label)
        count=len(item.periods[0].contributions);maximum=max(maximum,count)
        for c,v in enumerate((label,item.item.key,cursor-trace_start+1,count),1):literal(data.cell(meta+i,c),v)
        previous={}
        for p in item.periods:
            for contribution in p.contributions:
                a=contribution.activity_id;current=contribution.weighted_fraction
                before=previous.get(a,Decimal(0));delta=current-before if current is not None and before is not None else None
                amount=item.item.amount
                money=amount*current if amount is not None and amount>=0 and current is not None else None
                values=(a,contribution.share,contribution.raw_requirement,contribution.derived_requirement,contribution.state,
                        before,current,delta,money,
                        f'Payment Input!{letter(contribution.source_column)}{contribution.source_row}' if contribution.source_column else 'No entered cell',
                        p.status,p.mapping_coverage,p.input_coverage)
                for c,v in enumerate(values,1):literal(data.cell(cursor,c),('[blank]' if c==3 else 'Unknown') if v is None else v)
                previous[a]=current;cursor+=1
    wb.defined_names.add(DefinedName('BP_Items',attr_text=f"'_BP_Data'!$A${meta}:$A${meta+len(leaves)-1}"))
    for i,p in enumerate(projection.periods):literal(data.cell(cursor+i,1),p)
    wb.defined_names.add(DefinedName('BP_Periods',attr_text=f"'_BP_Data'!$A${cursor}:$A${cursor+periods-1}"))
    known = known_projected_totals(projection)
    known_start = cursor + periods + 2
    for offset, values in enumerate((known.cumulative, known.period)):
        for p, value in enumerate(values):
            literal(data.cell(known_start + offset, p + 1), 'Unavailable' if value is None else value)
    for ws,base,modes in [(pct,0,('Cumulative %','Period %')),(amt,2,('Cumulative Amount','Period Amount'))]:
        ws['A1']='BOQ Payment — Project';ws['A1'].font=Font(size=18,bold=True,color='1F4E78');ws.row_dimensions[1].height=28;ws['A2']='Projected Requirements — not certified payment'
        ws['A3']='View';ws['B3']=modes[0];dropdown(ws,'B3','"'+','.join(modes)+'"')
        ws['A4']='Select view → F9 → Read / Print / Save. Re-export to recover report edits.'
        ws['F3']='Open Trace';ws['F3'].hyperlink=Hyperlink(ref="F3",location="'Trace'!A1");ws['F3'].style='Hyperlink'
        valid=f'OR($B$3="{modes[0]}",$B$3="{modes[1]}")'
        ws['G3']=f'=IF({valid},$B$3,"Invalid view")'
        if base == 2:
            ws.merge_cells('A5:E5')
            ws['A5'] = 'Known Projected Amount — not certified payment'
            ws['F5'] = known.status
            ws['F5'].alignment = Alignment(wrap_text=True, vertical='center')
            ws.row_dimensions[5].height = 42
            for p in range(periods):
                col = letter(p + 1)
                ref = f"INDEX('_BP_Data'!${col}${known_start}:${col}${known_start+1},IF($B$3=\"{modes[0]}\",1,2))"
                ws.cell(5, p + 7, f'=IF(NOT({valid}),"Invalid view",{ref})')
                ws.cell(5, p + 7).number_format = '#,##0.00;[Red]-#,##0.00;0.00'
            for cell in ws[5]:
                cell.fill = PatternFill('solid', fgColor=PROJECT_FILL)
                cell.font = Font(bold=True, color='000000')
        for c,h in enumerate(('WBS' ,'BOQ Item','Unit','Qty','BOQ Amount','Mapping / status')+projection.periods,1):literal(ws.cell(6,c),h)
        for p,source in enumerate(projection.period_sources):literal(ws.cell(7,7+p),source.payment_date);ws.cell(7,7+p).number_format='dd-mmm-yy'
        for i,row in enumerate(rows):
            check();r=8+i
            for c,v in enumerate((row.wbs,row.name,row.unit,row.quantity,row.amount,row.status),1):literal(ws.cell(r,c),v)
            ws.cell(r,2).alignment=Alignment(indent=min(row.level,15),vertical='center',wrap_text=True)
            ws.cell(r,5).number_format='#,##0.00';ws.cell(r,4).number_format='#,##0.00'
            ws.row_dimensions[r].outlineLevel=min(row.level,7)
            ws.row_dimensions[r].height=38
            ws.cell(r,6).alignment=Alignment(wrap_text=True,vertical='center')
            for p in range(periods):
                col=letter(p+1);x=letter(p+7)
                ref=f"INDEX('_BP_Data'!${col}$1:${col}${4*n},IF($B$3=\"{modes[0]}\",{base*n+i+1},{(base+1)*n+i+1}))"
                ws.cell(r,p+7,f'=IF(NOT({valid}),"Invalid view",IF({ref}="Unavailable","",{ref}))')
                ws.cell(r,p+7).number_format='0.0%;[Red]-0.0%;""' if base==0 else '#,##0.00;[Red]-#,##0.00;""'
            if row.group:
                colors=(PALETTE.wbs_level_1_fill,PALETTE.wbs_level_2_fill,PALETTE.wbs_level_3_fill,PALETTE.wbs_level_4_fill)
                for c in range(1,7+periods):
                    fill=PROJECT_FILL if row.level==0 else colors[min(max(row.level-1,0),3)] if c<=6 else '2F5597' if row.level==1 else '5B9BD5'
                    ws.cell(r,c).fill=PatternFill('solid',fgColor=fill)
                    ws.cell(r,c).font=Font(bold=True,color='FFFFFF' if c>6 and row.level>0 else '000000')
        if base==0:
            from openpyxl.styles.differential import DifferentialStyle
            from openpyxl.styles.numbers import NumberFormat
            for p in range(1,periods):
                col,prev=letter(p+7),letter(p+6)
                rule=FormulaRule(formula=[f'AND($B$3="Cumulative %",ISNUMBER({col}8),{col}8=1,ISNUMBER({prev}8),{prev}8=1)'])
                rule.dxf=DifferentialStyle(numFmt=NumberFormat(numFmtId=164,formatCode=';;;'))
                ws.conditional_formatting.add(f'{col}8:{col}{7+n}',rule)
        ws.freeze_panes='G8';ws.sheet_properties.outlinePr.summaryBelow=False
        ws.sheet_view.showGridLines=False;ws.print_title_rows='1:7';ws.print_title_cols='A:F'
        ws.print_area=f'A1:{letter(6+periods)}{7+n}';ws.sheet_properties.pageSetUpPr.fitToPage=False
        ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A3;ws.page_setup.scale=80
        for c in range(13,6+periods,7):ws.col_breaks.append(Break(id=c))
        for c,width in zip('ABCDEF',(16,52,10,14,22,31)):ws.column_dimensions[c].width=width
        for c in range(7,7+periods):ws.column_dimensions[letter(c)].width=15
        for cell in ws[6]:cell.fill=PatternFill('solid',fgColor=HEADER_FILL);cell.font=Font(color='FFFFFF',bold=True)
        ws.row_dimensions[6].height=30
    trace['A1']='BOQ Payment — Trace';trace['A2']='BOQ Item';trace['B2']=labels[0];dropdown(trace,'B2','=BP_Items')
    trace['A3']='Period';trace['B3']=projection.periods[0];dropdown(trace,'B3','=BP_Periods')
    trace['A4']='Select Item and Period, then F9. Navigation links do not select an item.'
    trace['A5']='Known contributions are not final totals for incomplete mappings.'
    trace['P2']=f'=IFERROR(MATCH(B2,BP_Items,0),0)';trace['P3']='=IFERROR(MATCH(B3,BP_Periods,0),0)'
    trace['P4']=f'=IF(P2=0,0,INDEX(_BP_Data!D{meta}:D{meta+len(leaves)-1},P2))'
    trace['P5']=f'=IF(P2=0,0,INDEX(_BP_Data!C{meta}:C{meta+len(leaves)-1},P2))'
    trace['A6']='=IF(OR(P2=0,P3=0),"Invalid selector",IF(P4=0,"Unmapped — no explicit allocations",B3))'
    headers=('Activity ID','Mapping Share','Raw Requirement','Derived Requirement','State','Previous contribution','Current contribution','Signed change','Known Amount','Source cell','Status','Mapping coverage','Input coverage')
    for c,h in enumerate(headers,1):trace.cell(8,c,h)
    for k in range(maximum):
        for c in range(1,14):
            lookup=f"INDEX(_BP_Data!$A${trace_start}:$M${max(cursor-1,trace_start)},$P$5+($P$3-1)*$P$4+{k},{c})"
            trace.cell(9+k,c,f'=IF(OR($P$2=0,$P$3=0),"Invalid selector",IF({k}>=$P$4,"",IF({lookup}="[blank]","",{lookup})))')
            if c in (2,3,4,6,7,8,12,13):trace.cell(9+k,c).number_format='0.0%;[Red]-0.0%;0.0%'
            if c==9:trace.cell(9+k,c).number_format='#,##0.00;[Red]-#,##0.00;0.00'
    line=maximum+12;trace.cell(line,1,'Source lineage / diagnostics')
    for key,value in projection.provenance:
        line+=1;literal(trace.cell(line,1),key);literal(trace.cell(line,2),value)
    for d in projection.diagnostics:
        line+=1
        for c,v in enumerate((d.code,d.identity,d.detail),1):literal(trace.cell(line,c),v)
    trace.column_dimensions['P'].hidden=True;trace.column_dimensions['A'].width=28;trace.column_dimensions['B'].width=65
    for c in range(3,14):trace.column_dimensions[letter(c)].width=24
    trace.freeze_panes='C9';trace.sheet_view.showGridLines=False
    for cell in trace[8]:cell.fill=PatternFill('solid',fgColor=HEADER_FILL);cell.font=Font(color='FFFFFF',bold=True)
    trace.row_dimensions[8].height=36
    request_initial_manual_excel_recalculation(wb)
    return wb
