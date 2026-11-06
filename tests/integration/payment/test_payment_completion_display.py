"""Completion suppression must remain formatting, never a numerical rewrite."""
import pytest
from openpyxl import load_workbook, Workbook
from scripts.payment_completion_proof import create_proof, VALUES
from tests.integration.payment.test_payment_requirement_summary import recalculate
from progress_studio.infrastructure.excel.payment_input_summary import add_amount_summaries, WBS_PERCENT_FORMAT


def test_completion_format_and_native_numbers(tmp_path):
    path=create_proof(tmp_path/'completion.xlsx')
    wb=load_workbook(path);ws=wb.active
    rules=[(str(cf.sqref),r) for cf in ws.conditional_formatting for r in ws.conditional_formatting[cf]
           if r.dxf and r.dxf.numFmt]
    assert len(rules)==2
    assert rules[0][1].dxf.fill.fgColor.rgb == "00FFF2CC"
    target,rule=rules[1]
    assert target=='G8:Q8'
    assert rule.formula==['IFERROR(AND(ISNUMBER(G8),G8=1,ISNUMBER(F8),F8=1),FALSE)']
    assert rule.dxf.numFmt.formatCode==';;;' and rule.stopIfTrue is False
    assert rule.dxf.fill is None
    assert ws['F8'].number_format==WBS_PERCENT_FORMAT
    assert all(ws.cell(9,i+6).value==v for i,v in enumerate(VALUES))
    assert all(ws.cell(9,i+6).number_format=='0.0%' for i in range(len(VALUES)))
    # Errors, text and unknowns cannot activate the numeric-only condition;
    # near-100 and previous near-100 cannot activate exact equality.
    assert 'ROUND' not in rule.formula[0]
    wb.close()
    calculated=recalculate(path,tmp_path/'calculated')
    wb=load_workbook(calculated,data_only=True);ws=wb.active
    assert [ws.cell(8,i+6).value for i in range(len(VALUES))]==pytest.approx(VALUES)
    assert [ws.cell(9,i+6).value for i in range(len(VALUES))]==pytest.approx(VALUES)
    assert 'Review decreasing' in ws['D4'].value
    wb.close()


@pytest.mark.parametrize('periods',[1,2,120])
def test_completion_rule_bounds(periods):
    wb=Workbook();ws=wb.active
    add_amount_summaries(ws,[dict(row_type='WBS',outline_level=1),
        dict(row_type='ACT',outline_level=2,amount=0)],periods)
    rules=[r for cf in ws.conditional_formatting for r in ws.conditional_formatting[cf]
           if r.dxf and r.dxf.numFmt]
    assert len(rules)==2*int(periods>1)
    assert ws['F8'].data_type=='f'
