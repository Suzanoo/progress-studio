"""Reproducible Payment Input proof with nested and repeated WBS branches."""
from datetime import datetime
from openpyxl import load_workbook
from tests.integration.payment.test_payment_line_renderer import _progress_workbook


def requirement_source(path):
    _progress_workbook(path)
    wb = load_workbook(path)
    ws = wb['main']
    ws.delete_rows(6, ws.max_row - 5)
    rows = [
        ('WBS', '1', 'Repeated branch', '', 1, None),
        ('Activity', '1.1', 'Foundation', 'A100', 2, 100),
        ('WBS', '1.2', 'Nested work', '', 2, None),
        ('Activity', '1.2.1', 'Structure', 'A200', 3, 300),
        ('Activity', '1.2.2', 'Milestone (zero weight)', 'M1', 3, 0),
        ('WBS', '2', 'Repeated branch', '', 1, None),
        ('Activity', '2.1', 'Independent branch', 'A300', 2, 200),
        ('WBS', '3', 'Unavailable weight', '', 1, None),
        ('Activity', '3.1', 'Missing Amount example', 'A400', 2, None),
        ('WBS', '4', 'Zero weight', '', 1, None),
        ('Activity', '4.1', 'Zero Amount example', 'A500', 2, 0),
        ('WBS', '5', 'Empty branch', '', 1, None),
    ]
    for r, (kind, wbs, name, identity, level, amount) in enumerate(rows, 6):
        for col, val in {1:kind, 2:wbs, 3:name, 4:'P', 5:identity, 8:level,
                         9:datetime(2026,3,2),10:datetime(2026,4,3),15:amount}.items():
            ws.cell(r,col,val)
        if kind == 'Activity':
            for col, val in zip(range(18,23), [.1,.15,.2,.2,.35]):
                ws.cell(r,col,val)
    # Actual duplicate ID is deliberately ignored by both Amount adapters.
    ws.cell(18,1,'Activity');ws.cell(18,4,'A');ws.cell(18,5,'A100');ws.cell(18,15,99999)
    wb.create_sheet('User Notes')['A1'] = 'PRESERVE'
    wb.save(path);wb.close()
    return path


def edit_requirements(path):
    wb=load_workbook(path);ws=wb['Payment Input']
    values={'A100':(1,0,None),'A200':(.2,None,None),'M1':(1,0,None),
            'A300':(.5,None,None),'A400':(.5,None,None),'A500':(0,None,None)}
    for r in range(8,ws.max_row+1):
        identity=ws.cell(r,3).value
        if identity in values:
            for c,v in enumerate(values[identity],6):ws.cell(r,c).value=v
    wb.save(path);wb.close()
