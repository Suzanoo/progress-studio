"""Portable, source-bound payment plan editing and validated readback."""
from dataclasses import dataclass
from decimal import Decimal, localcontext
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree
from zipfile import ZipFile
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from progress_studio.domain.payment_plan_adjustment import adjusted_selection
from progress_studio.infrastructure.excel.payment_workbook import MAIN_NS, PaymentWorkbookSnapshotter

NAMES = ('Edit plan','Summary','Period percent','Period amount','Cumulative amount','Original','Identity')


def template(source, sel):
    if not sel.rows or sel.revision != source.revision:
        raise ValueError('Select work from the current payment workbook.')
    w = Workbook()
    w.remove(w.active)
    for name in NAMES:
        w.create_sheet(name)
    identity = w['Identity']
    for item in [('Format','Payment Plan Adjustment v1'),('Project',source.project),('Revision',source.revision),
                 ('Monetary source',source.monetary_origin),('Selected basis',str(sel.selected.basis)),
                 ('Project basis',str(sel.project.basis)),('Stable key','Selected ancestor')]:
        identity.append(item)
    for row in sel.rows:
        identity.append([row.key,sel.included_in.get(row.key)])
    last = 3+len(sel.rows)
    for name in (NAMES[0],*NAMES[2:6]):
        s=w[name]
        s.append([name+' | '+source.project,'Edit blue cumulative percentages' if name=='Edit plan' else 'Source-bound calculation'])
        s.append(['Reporting row','Description','Role','Amount basis',*[p.id for p in source.periods]])
        s.append([None,None,None,'Eligible date',*[p.eligible for p in source.periods]])
        for n,row in enumerate(sel.rows,4):
            scope=sel.row_values[row.key]
            s.append([row.activities[0] if row.activity else row.code or 'Project',row.name,
                      'Included' if row.key in sel.included_in else 'Counted',scope.basis])
            for j,v in enumerate(scope.values,5):
                c=s.cell(n,j); ref=c.coordinate
                prev=get_column_letter(j-1)+str(n)
                grid='Cumulative amount' if name=='Period amount' else 'Edit plan'
                a=f"'{grid}'!{ref}"; b=f"'{grid}'!{prev}"
                if name in ('Original','Edit plan'):
                    c.value=float(v.fraction) if v.fraction is not None else None
                elif name=='Cumulative amount':
                    c.value = (f'=IF(COUNT({a},$D{n})=2,{a}*$D{n},"")' if v.money is not None else '=""')
                elif j==5:
                    c.value=f'=IF(ISNUMBER({a}),{a},"")'
                else:
                    c.value=f'=IF(COUNT({a},{b})=2,{a}-{b},"")'
                if name=='Edit plan':
                    c.protection=Protection(locked=False)
                    c.fill=PatternFill('solid',fgColor='DAEEFC')
                c.number_format='#,##0.00' if 'amount' in name else '0.00%'
        s.freeze_panes='E4'
        for col,width in [('A',29),('B',52),('C',14),('D',20)]:
            s.column_dimensions[col].width=width
    end=get_column_letter(4+len(source.periods))
    dv=DataValidation(type='decimal',operator='between',formula1=0,formula2=1,allow_blank=True,showErrorMessage=True)
    w['Edit plan'].add_data_validation(dv);dv.add(f'E4:{end}{last}')
    s=w['Summary'];s.append(['Period','Eligible date','Original period','Adjusted period','Original cumulative','Adjusted cumulative'])
    for j,p in enumerate(source.periods):
        v=sel.selected.values[j];s.append([p.id,p.eligible,v.movement,None,v.money])
        for col,grid in [(4,'Period amount'),(6,'Cumulative amount')]:
            letter=get_column_letter(5+j);r=f"'{grid}'!{letter}4:{letter}{last}";role=f"'{grid}'!C4:C{last}"
            s.cell(j+2,col,f'=IF(SUMPRODUCT(--({role}="Counted"),--ISNUMBER({r}))=COUNTIF({role},"Counted"),SUMIF({role},"Counted",{r}),"")')
        for c in s[j+2][2:]:c.number_format='#,##0.00'
    for sh in w:
        sh.protection.sheet=True
        for col in range(1,sh.max_column+1):
            key=get_column_letter(col)
            if key not in sh.column_dimensions:sh.column_dimensions[key].width=22
        for cell in sh[2 if sh.title in ('Edit plan','Original',*NAMES[2:5]) else 1]:
            cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='275C60')
    identity.column_dimensions['B'].width=72
    # Project labels and source keys are literal workbook data.
    for sh in w:
        for row in sh:
            for cell in row:
                if cell.data_type == 'f' and (sh.title in ('Edit plan','Original','Identity') or
                        sh.title in NAMES[2:5] and (cell.row < 4 or cell.column < 5)):
                    cell.data_type='s'
    return w


def create_adjustment(path,source,selection):
    target=Path(path)
    if target.suffix.lower()!='.xlsx':raise ValueError('Choose an .xlsx filename.')
    with NamedTemporaryFile(dir=target.parent,suffix='.xlsx',delete=False) as f:tmp=Path(f.name)
    try:
        template(source,selection).save(tmp)
        os.link(tmp,target)
    finally:tmp.unlink(missing_ok=True)


@dataclass(frozen=True)
class Reviewed:
    path: Path
    revision: str
    selection: object
    issues: tuple
    incomplete: bool

    def validate(self):
        if self.incomplete:raise ValueError('Resolve missing requirements before adjusted export.')
        if sha256(self.path.read_bytes()).hexdigest()!=self.revision:raise ValueError('Adjustment workbook changed. Reopen and review it.')


def read_adjustment(path,source):
    path=Path(path);raw=path.read_bytes()
    with ZipFile(BytesIO(raw)) as z:
        if any('externalLinks/' in n or 'vbaProject' in n for n in z.namelist()):
            raise ValueError('Use the original macro-free adjustment workbook.')
        parts = PaymentWorkbookSnapshotter._sheet_map(z)
        if tuple(parts) != NAMES:
            raise ValueError('Unexpected adjustment workbook layout. Create a new adjustment workbook.')
        def literals(sheet):
            root = ElementTree.fromstring(z.read(parts[sheet]))
            ns = '{' + MAIN_NS + '}'
            return {c.get('r'): Decimal(c.findtext(ns+'v')) for c in root.iter(ns+'c')
                    if c.get('t','n')=='n' and c.findtext(ns+'v') and c.find(ns+'f') is None}
        inputs = literals('Edit plan')
        originals = literals('Original')
    w=load_workbook(BytesIO(raw))
    if tuple(w.sheetnames)!=NAMES or w['Identity']['B3'].value!=source.revision:
        raise ValueError('Open the matching Payment Workbook and its adjustment file.')
    keys=[c.value for (c,) in w['Identity'].iter_rows(min_row=8,max_col=1)]
    sel=source.select(keys,revision=source.revision)
    expected=template(source,sel);edits={};issues=[]
    for s in expected:
        actual=w[s.title]
        if (s.max_row,s.max_column)!=(actual.max_row,actual.max_column):raise ValueError(f'Unexpected dimensions in {s.title}.')
        for row in s:
            for c in row:
                x=actual[c.coordinate].value;t=c.value
                if s.title=='Edit plan' and c.row>=4 and c.column>=5:
                    if x is not None and (type(x) not in (int,float) or not 0<=x<=1):
                        raise ValueError(f'Enter a numeric 0–100% or clear {s.title}!{c.coordinate}.')
                    if inputs.get(c.coordinate)!=originals.get(c.coordinate):
                        edits.setdefault(keys[c.row-4],{})[c.column-5]=inputs.get(c.coordinate)
                    continue
                if isinstance(t,(Decimal,float)) and type(x) in (float,int) and abs(float(t)-x)<=abs(float(t))*1e-14:continue
                if hasattr(x,'date'):x=x.date()
                if x!=t:raise ValueError(f'Unexpected {s.title}!{c.coordinate}. Recreate the adjustment workbook.')
    plan=adjusted_selection(sel,edits)
    incomplete=False
    for row in plan.rows:
        for j,v in enumerate(plan.row_values[row.key].values):
            if v.money is None or v.fraction is None:
                incomplete=True;issues.append(f'{row.code or row.key} / {source.periods[j].id}: unresolved')
    for parent in set(plan.included_in.values()):
        children=[plan.row_values[k] for k,p in plan.included_in.items() if p==parent]
        for j,v in enumerate(plan.row_values[parent].values):
            with localcontext() as ctx:
                ctx.prec=80
                known=sum((c.values[j].money or 0) for c in children)
            same_scope = sum(len(c.activities) for c in children)==len(plan.row_values[parent].activities)
            if v.money is not None and (known>v.money or (same_scope and all(c.values[j].money is not None for c in children) and known!=v.money)):
                issues.append(f'{parent} / {source.periods[j].id}: included detail is inconsistent with parent')
    return Reviewed(path,sha256(raw).hexdigest(),plan,tuple(issues),incomplete)
