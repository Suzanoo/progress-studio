"""Optional independent formula gate: pip install formulas (test-only).

The evaluator lacks ISFORMULA and has incomplete defined-name resolution for
blank cells. Inline names to their exact addresses and resolve ISFORMULA from
actual input cell metadata in the disposable evaluator copy only. Financial and
validation arithmetic is left intact. This does not certify Desktop Excel.
"""
import re
from datetime import date
import pytest
from openpyxl.utils.datetime import to_excel
from tests.unit.finance.test_finance_workbook import sample
from progress_studio.infrastructure.excel.finance_input_workbook import read_finance_inputs
from progress_studio.infrastructure.excel.financial_forecast_workbook import build_finance_view
from progress_studio.services.financial_forecast_deriver import FinancialForecastDeriver

formulas=pytest.importorskip('formulas',reason='Optional Excel formula evaluator is not installed')


def evaluated(wb,tmp_path):
    names={n:d.attr_text for n,d in wb.defined_names.items()}
    for ws in wb:
        for row in ws:
            for cell in row:
                if cell.data_type!='f':continue
                def isformula(match):
                    sheet,addr=match[1].rsplit('!',1)
                    return 'TRUE' if wb[sheet.strip("'")][addr].data_type=='f' else 'FALSE'
                f=re.sub(r'_xlfn.ISFORMULA\(([^)]+)\)',isformula,cell.value)
                for name,target in names.items():f=re.sub(r'\b'+re.escape(name)+r'\b',target,f)
                cell.value=f
    path=tmp_path/'eval.xlsx';wb.save(path)
    solution=formulas.ExcelModel().loads(str(path)).finish().calculate()
    def get(sheet,cell):
        key=next(k for k in solution if k.endswith(sheet.upper()+"'!"+cell))
        return solution[key].value[0,0]
    return get


CASES=['base','zero','amount_override','date_override','clear_date','snapshot','future_actual',
       'advance_cutoff','financing','opening_later','opening_deficit','incomplete','overdue',
       'undated','rounding','credit_days','correction','same_day','later_release','text_booleans']


@pytest.mark.parametrize('case',CASES)
def test_live_finance_matches_ff1_after_input_edit(case,tmp_path):
    w=sample();s=w['Finance Input'];build_finance_view(w,monthly_capacity=24)
    # All edits occur after formulas exist; no rebuild before evaluation.
    if case in ('zero','amount_override','date_override','clear_date'):
        s['U25']='cert:net';s['Y25']='Scenario'
        if case=='zero':s['V25']=0
        if case=='amount_override':s['V25']=75
        if case=='date_override':s['W25']=date(2026,4,10)
        if case=='clear_date':s['X25']=True
    elif case=='snapshot':s['K26']=80;s['O26']=date(2026,1,15);s['F25']='invoice';s['C25']=20
    elif case=='future_actual':s['B25']=date(2026,2,15)
    elif case=='advance_cutoff':s['B7']=date(2026,3,15)
    elif case=='financing':s['N25']='financing'
    elif case=='opening_later':s['B5']=date(2026,1,25)
    elif case=='opening_deficit':s['B6']=-500
    elif case=='incomplete':s['B10']=False;s['B11']=False;s['B12']=False
    elif case=='overdue':s['M25']=date(2026,1,10)
    elif case=='undated':s['M26']=None
    elif case=='rounding':s['K27']=100.01;s['B9']=0.075
    elif case=='credit_days':s['B8']=31
    elif case=='correction':
        for c,value in enumerate(['refund',date(2026,1,25),-10,'in','operating','cert:net','Refund'],1):s.cell(26,c,value)
    elif case=='same_day':s['M26']=date(2026,2,5)
    elif case=='later_release':s['R27']=date(2027,7,1);s['B14']=date(2026,3,1)
    elif case=='text_booleans':
        for r in range(10,14):s.cell(r,2,'TRUE')
    expected=FinancialForecastDeriver().derive(read_finance_inputs(w));get=evaluated(w,tmp_path)
    assert get('Cash Flow','B5')==0
    assert get('Cash Flow','B4')==('COMPLETE' if expected.forecast_complete else 'CONDITIONAL / INCOMPLETE')
    for cell,value in [('B7',expected.peak_funding_requirement),('B9',expected.minimum_balance),
                       ('B10',expected.remaining_receivable),('B11',expected.remaining_cash_out),
                       ('B12',expected.remaining_retention),('B13',expected.operating_cash_out_ac_proxy)]:
        assert float(get('Cash Flow',cell))==pytest.approx(float(value),abs=0.000001)
    assert get('Cash Flow','B8')==to_excel(expected.minimum_balance_date)
    for row,point in enumerate(expected.monthly_points,19):
        for c,value in zip('BCDEF',(point.actual_in,point.actual_out,point.forecast_in,point.forecast_out,point.closing_balance)):
            assert float(get('Cash Flow',f'{c}{row}'))==pytest.approx(float(value),abs=0.000001)


@pytest.mark.parametrize('cell,value', [('K25','text'),('J25','bad*id'),('B8',1.5),('B9',1.1),
    ('B10','yes'),('C25',-1),('F25','unknown'),('V25',-1),('W25',date(2026,3,1)),
    ('K25','=100'),('B6','=10'),('B14',date(2030,1,1))])
def test_live_invalid_edit_blocks_headline(cell,value,tmp_path):
    w=sample();build_finance_view(w,monthly_capacity=24);s=w['Finance Input'];s[cell]=value
    get=evaluated(w,tmp_path)
    assert get('Cash Flow','B5')>0
    assert get('Cash Flow','B4')=='INVALID INPUT'
    assert str(get('Cash Flow','B7'))=='#N/A'
