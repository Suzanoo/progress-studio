"""Small disposable formula proof; does not claim native Windows Excel acceptance."""
import shutil,subprocess
from openpyxl import load_workbook
import pytest
from progress_studio.infrastructure.excel.boq_payment_workbook import render
from tests.integration.payment.test_boq_mapping_trace_workbook import example,values


def test_all_trace_modes_recalculate(tmp_path):
    office=shutil.which('soffice') or shutil.which('libreoffice')
    if not office:pytest.skip('LibreOffice unavailable; native Excel remains a separate gate')
    incoming=tmp_path/'incoming';out=tmp_path/'out';incoming.mkdir();out.mkdir()
    cases=('boq','reverse','unmapped','manual_mapping','invalid_boq','invalid_activity')
    for case in cases:
        w=render(example());labels=[row[0] for row in values(w,'BP_Items')]
        if case=='reverse':
            w['BOQ Mapping']['B2']='INVALID BOQ';w['Activity Mapping']['B2']='A | Duplicate name'
        if case=='unmapped':w['Activity Mapping']['B2']='C | No mapping'
        if case=='manual_mapping':w['BOQ Mapping']['B2']=labels[2]
        if case=='invalid_boq':w['BOQ Mapping']['B2']='INVALID BOQ'
        if case=='invalid_activity':w['Activity Mapping']['B2']='INVALID ACTIVITY'
        w.calculation.calcMode='auto';w.save(incoming/f'{case}.xlsx');w.close()
    command=[office,'-env:UserInstallation='+ (tmp_path/'profile').as_uri(),'--headless','--convert-to','xlsx','--outdir',str(out)]
    result=subprocess.run(command+[str(incoming/f'{case}.xlsx') for case in cases],capture_output=True,text=True,timeout=180)
    assert result.returncode==0,result.stdout+result.stderr
    for case in cases:
        w=load_workbook(out/f'{case}.xlsx',data_only=True)
        name='Activity Mapping' if case in ('reverse','unmapped','invalid_activity') else 'BOQ Mapping'
        t=w[name];row=tuple(t.cell(9,c).value for c in range(1,14))
        if case=='boq':
            assert row[:4]==('A','Duplicate name',.6,60)
            assert (t['B7'].value,t['D7'].value)==(1,100)
            assert tuple(t.cell(10,c).value for c in range(1,4))==('B','Duplicate name',.4)
        elif case=='reverse':
            assert row[:4]==('a','BOQ a',.6,60)
            assert t['D7'].value==85 and t['B7'].value is None
            assert tuple(t.cell(10,c).value for c in range(1,4))==('b','BOQ b',.25)
        elif case in ('unmapped','manual_mapping'):
            assert row[0] in (None,'') and ('No saved Mapping' in t['A6'].value or 'No Activity Mapping' in t['A6'].value)
        else:assert t['A6'].value=='Invalid selector' and row[0]=='Invalid selector'
        w.close()
