"""Small deterministic OOXML/P6 fixtures, including explicit saved formula caches."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

P6_NS = 'http://xmlns.oracle.com/Primavera/P6Professional/V24.12/API/BusinessObjects'
SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def sheet(rows):
    return f'<worksheet xmlns="{SS_NS}"><sheetData>' + ''.join(
        f'<row r="{r}">' + ''.join(cells) + '</row>' for r, cells in rows) + '</sheetData></worksheet>'


def text_cell(ref, value):
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def number_cell(ref, value, *, formula=None, kind=None):
    attr = f' t="{kind}"' if kind else ''
    f = f'<f>{escape(formula)}</f>' if formula is not None else ''
    v = '' if value is None else f'<v>{escape(str(value))}</v>'
    return f'<c r="{ref}"{attr}>{f}{v}</c>'


def workbook(path, amounts=('12.34567890123456789', '0', '0'), *, ids=('A1','A2','M1'),
             flags=(False,False,True), formulas=None, kinds=None, shifted=False):
    formulas = formulas or (None,) * len(ids)
    kinds = kinds or (None,) * len(ids)
    amount_col = 'H' if shifted else 'D'
    rows = [(4, [text_cell('A4','Row Type'), text_cell('B4','P/A'), text_cell('C4','Activity ID'),
                 text_cell(amount_col+'4','Amount'), text_cell('E4','Description')])]
    for i, (aid, value, formula, kind) in enumerate(zip(ids, amounts, formulas, kinds)):
        n = 5 + 2*i
        rows.append((n, [text_cell(f'A{n}','Activity'), text_cell(f'B{n}','P'), text_cell(f'C{n}',aid),
                         number_cell(f'{amount_col}{n}',value,formula=formula,kind=kind),text_cell(f'E{n}','Work '+aid)]))
        rows.append((n+1, [text_cell(f'A{n+1}','Activity'), text_cell(f'B{n+1}','A'),text_cell(f'C{n+1}',aid),
                           number_cell(f'{amount_col}{n+1}','999999')]))
    sheets = {'main':sheet(rows)}
    if flags is not None:
        ev = [(3,[text_cell('A3','Activity ID'),text_cell('B3','Milestone')])]
        for i,(aid,flag) in enumerate(zip(ids,flags),4):
            ev.append((i,[text_cell(f'A{i}',aid),number_cell(f'B{i}',int(flag),kind='b')]))
        sheets['EV Monetary Inputs'] = sheet(ev)
    write_package(path,sheets)
    return Path(path)


def write_package(path, sheets):
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{SS_NS}" xmlns:r="{rel}"><sheets>'+''.join(
            f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i,name in enumerate(sheets,1))+'</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(
            f'<Relationship Id="rId{i}" Target="worksheets/sheet{i}.xml" Type="{rel}/worksheet"/>' for i in range(1,len(sheets)+1))+'</Relationships>')
        for i,xml in enumerate(sheets.values(),1):
            z.writestr(f'xl/worksheets/sheet{i}.xml',xml)


def edit_package(path, member, transform):
    with ZipFile(path) as z:
        entries = [(n,z.read(n)) for n in z.namelist()]
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for n,data in entries:
            z.writestr(n,transform(data.decode()).encode() if n==member else data)


def p6(path, *, declared=True, prefix=False):
    definition = '<UDFType><DataType>Cost</DataType><ObjectId>47</ObjectId><SubjectArea>Activity</SubjectArea><Title>Budget</Title></UDFType>' if declared else ''
    activities = []
    for i,(aid,value,kind) in enumerate((('A1','1','Task Dependent'),('A2','9','Resource Dependent'),
                                        ('M1','500.50','Finish Milestone'),('NKC2-A1','888','Task Dependent')),10):
        udf = f'<UDF><TypeObjectId>47</TypeObjectId><CostValue>{value}</CostValue></UDF>' if declared else ''
        activities.append(f'''<Activity><Id>{aid}</Id><ObjectId>{i}</ObjectId><Name>Work &amp; {aid}</Name><Type>{kind}</Type>
<WBSObjectId>2</WBSObjectId><PlannedStartDate>2026-01-01T08:00:00</PlannedStartDate><PlannedFinishDate>2026-01-02T17:00:00</PlannedFinishDate><PlannedDuration>8</PlannedDuration>
<BudgetedTotalCost>12345</BudgetedTotalCost><x:Extra flag="a &gt; b">unchanged</x:Extra>{udf}</Activity>''')
    xml = f'''<?xml version="1.0" encoding="utf-8"?>
<APIBusinessObjects xmlns="{P6_NS}" xmlns:x="urn:external" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<?keep unchanged?><!-- preserve comment -->{definition}<Currency><ObjectId>300</ObjectId></Currency>
<Calendar><ObjectId>20</ObjectId><Name>Calendar</Name></Calendar>
<Project><Id>Combined</Id><Name>Combined project</Name><ObjectId>1</ObjectId>
<WBS><ObjectId>2</ObjectId><Code>1</Code><Name>Work</Name></WBS>
{''.join(activities)}<Relationship><PredecessorActivityObjectId>10</PredecessorActivityObjectId><SuccessorActivityObjectId>11</SuccessorActivityObjectId><Type>Finish to Start</Type></Relationship>
<ResourceAssignment><ObjectId>99</ObjectId><PlannedCost>999</PlannedCost></ResourceAssignment></Project></APIBusinessObjects>'''
    if prefix:
        import re
        xml = xml.replace(f'xmlns="{P6_NS}"',f'xmlns:p="{P6_NS}"')
        xml = re.sub(r'<(/?)([A-Z][A-Za-z0-9]*)(?=[\s/>])',r'<\1p:\2',xml)
    Path(path).write_text(xml,encoding='utf-8')
    return Path(path)
