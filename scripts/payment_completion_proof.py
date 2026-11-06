"""Generate a small editable proof of WBS completion formatting (no engine changes)."""
import argparse
import json
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from progress_studio.infrastructure.excel.payment_input_summary import add_amount_summaries

VALUES = [.94, 1, 1, .95, 1, 1, .9999, 1, 1, 0, 1e-9, 1]

def create_proof(path):
    wb = Workbook(); ws = wb.active; ws.title = 'Payment Input'
    for col, value in enumerate(['Type','WBS','Activity ID','Activity Name','Amount'],1):
        ws.cell(6,col,value)
    for i in range(len(VALUES)): ws.cell(6,i+6,f'P{i+1:02}')
    tree = [dict(row_type='WBS',outline_level=1),dict(row_type='ACT',outline_level=2,amount=100)]
    for row, kind, name in [(8,'WBS','Completion proof'),(9,'ACT','Editable source')]:
        ws.cell(row,1,kind);ws.cell(row,4,name)
    ws['C9']='PROOF-1'
    for i,value in enumerate(VALUES):
        ws.cell(9,i+6,value).number_format='0.0%'
    add_amount_summaries(ws,tree,len(VALUES))
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=1
    wb.calculation.calcMode='auto';wb.calculation.fullCalcOnLoad=True
    wb.save(path)
    expected={'numeric_WBS':VALUES,'hidden_periods':['P03','P06','P09','P10'],
              'visible_near_100':'P07', 'visible_small_nonzero':'P11',
              'visible_returns':['P05','P08','P12'],
              'note':'P10 uses existing exact-zero format; all Activity values visible. Amber warning persists.'}
    Path(path).with_suffix('.json').write_text(json.dumps(expected,indent=2),encoding='utf-8')
    return path

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('output',type=Path)
    args=parser.parse_args();args.output.parent.mkdir(parents=True,exist_ok=True)
    create_proof(args.output)
