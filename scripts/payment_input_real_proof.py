"""PI-1 real-workbook verification, always on copies; optional native recalculation.

python scripts/payment_input_real_proof.py SOURCE.xlsx --output tmp/pi1-real --recalculate
"""
from pathlib import Path
import argparse
from dataclasses import asdict
import hashlib
import json
import math
import re
import shutil
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from progress_studio.services.payment_service import PaymentService
from progress_studio.services.rebuild_service import WorkbookRebuildEngine


def sparse(service, path):
    data = service.read_payment_requirements(path)
    return {p.period_id: {r.activity_id: r.required_fraction for r in p.requirements} for p in data.periods}


def run(source, output, recalculate=False):
    source = source.resolve(); output = output.resolve(); output.mkdir(parents=True, exist_ok=True)
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    service = PaymentService(); before = sparse(service, source)
    original_positions = service.prepare_payment_positions(source, source)
    corrected = output/'NKC2_CUMULATIVE_INPUT.xlsx'
    if corrected == source:
        raise ValueError('Proof output must not overwrite source')
    service.prepare_embedded_payment_input(source, corrected)
    assert sparse(service, corrected) == before
    assert service.prepare_payment_positions(corrected, corrected) == original_positions
    wb = load_workbook(corrected); ws = wb['Payment Input']
    periods = [c.column for c in ws[6] if re.fullmatch(r'P\d+', str(c.value))]
    acts = [r for r in range(8, ws.max_row+1) if ws.cell(r,1).value == 'ACT']
    parents = [r for r in range(8, ws.max_row+1) if ws.cell(r,1).value == 'WBS']
    assert not any(ws.cell(r,1).value == 'COV' for r in range(8, ws.max_row+1))
    assert parents[0] == 8 and ws.row_dimensions[8].outlineLevel == 0
    assert all(ws.row_dimensions[r].outlineLevel > 0 for r in range(9, ws.max_row+1))
    assert ws.cell(8,3).value is None
    project_amount = sum(ws.cell(r,5).value for r in acts)
    assert math.isclose(ws.cell(8,5).value, project_amount, rel_tol=1e-12)
    expected = []
    for row in parents:
        end = next((r for r in parents if r > row and ws.row_dimensions[r].outlineLevel <= ws.row_dimensions[row].outlineLevel), ws.max_row+1)
        children = [r for r in acts if row < r < end]
        weights = [ws.cell(r,5).value for r in children]
        assert all(isinstance(x, (int,float)) for x in weights), 'This numerical proof requires valid real-source weights'
        total = sum(weights)
        schedules = {r: [ws.cell(r,c).value for c in periods] for r in children}
        missing = any(ws.cell(r,5).value > 0 and all(v is None for v in schedules[r]) for r in children)
        populated = any(v is not None for seq in schedules.values() for v in seq)
        states = dict.fromkeys(children, 0.0); values = []
        for index in range(len(periods)):
            for r in children:
                if schedules[r][index] is not None: states[r] = float(schedules[r][index])
            value = (None if not populated else 'No weight' if total == 0 else 'Incomplete' if missing
                     else sum(ws.cell(r,5).value*states[r] for r in children)/total)
            values.append(value)
        expected.append({'wbs':ws.cell(row,2).value, 'row':row, 'values':values})
    oracle = output/'EXPECTED_VALUES.json'; oracle.write_text(json.dumps(expected,ensure_ascii=False,indent=2),encoding='utf-8')
    evidence = {'source_sha256':original_hash,'preserved_requirements':sum(map(len,before.values())),
                'activity_count':len(acts),'wbs_count':len(parents),'periods':len(periods),
                'positions_preserved_after_prepare':True, 'project_summary_name':ws.cell(8,4).value,
                'project_amount':project_amount, 'project_summary_count':1}
    if recalculate:
        office = shutil.which('soffice') or shutil.which('libreoffice')
        if not office:raise RuntimeError('Native LibreOffice not available')
        incoming=output/'calc-incoming'; incoming.mkdir(exist_ok=True)
        calcdir=output/'calculated';calcdir.mkdir(exist_ok=True)
        # Isolate presentation formulas; unrelated source charts/calculation are not transformed.
        for sheet in list(wb):
            if sheet.title != 'Payment Input':wb.remove(sheet)
        wb.calculation.calcMode='auto';wb.calculation.fullCalcOnLoad=True;wb.calculation.forceFullCalc=True
        disposable=incoming/'CUMULATIVE_DISPLAY.xlsx';wb.save(disposable)
        with (output/'recalculation.log').open('w',encoding='utf-8') as log:
            result=subprocess.run([office,'-env:UserInstallation='+(output/'office-profile').as_uri(),
                '--headless','--convert-to','xlsx','--outdir',str(calcdir),str(disposable)],
                stdout=log,stderr=subprocess.STDOUT,timeout=180)
        assert result.returncode==0
        cached=load_workbook(calcdir/disposable.name,data_only=True); sheet=cached['Payment Input']
        for item in expected:
            for col,value in zip(periods,item['values']):
                actual=sheet.cell(item['row'],col).value
                if isinstance(value,(int,float)):assert math.isclose(actual,value,abs_tol=1e-10),(item,col,actual)
                else:assert (actual or None)==value,(item,col,actual)
        evidence['calculated_values_verified']=len(expected)*len(periods);cached.close()
    wb.close()
    # Exercise existing current saved-input paths, including zero and clear.
    edited=output/'EDITED_INPUT.xlsx';wb=load_workbook(corrected);ws=wb['Payment Input']
    rows={ws.cell(r,3).value:r for r in acts}
    required={'NKC2-A1010','NKC2-A1050','NKC2-A1020'}
    if not required.issubset(rows):raise ValueError('This workflow proof uses the documented NKC2 fixture IDs')
    ws.cell(rows['NKC2-A1010'],8,.25)
    ws.cell(rows['NKC2-A1050'],9,0)
    ws.cell(rows['NKC2-A1020'],9).value=None
    wb.save(edited);wb.close();saved=sparse(service,edited)
    assert saved['P03']['NKC2-A1010']==.25 and saved['P04']['NKC2-A1050']==0 and 'NKC2-A1020' not in saved['P04']
    positions=service.prepare_payment_positions(edited,edited)
    old=next(p for p in original_positions.periods[2].points if p.activity_id=='NKC2-A1010')
    new=next(p for p in positions.periods[2].points if p.activity_id=='NKC2-A1010')
    assert old.timescale_column!=new.timescale_column
    evidence['changed_point']={'before':asdict(old),'after':asdict(new)}
    engine=WorkbookRebuildEngine()
    for name,method in [('snapshot',engine.rebuild_payment),('live',engine.rebuild_live_payment)]:
        target=output/f'{name.upper()}_PAYMENT_REBUILT.xlsx';result=method(edited,target)
        assert sparse(service,target)==saved
        assert service.prepare_payment_positions(target,target)==positions
        a=load_workbook(edited);b=load_workbook(target)
        assert list(a['Payment Input'].values)==list(b['Payment Input'].values)
        a.close();b.close();evidence[name]=asdict(result)
    assert hashlib.sha256(source.read_bytes()).hexdigest()==original_hash
    evidence['source_unchanged']=True
    (output/'VERIFICATION.json').write_text(json.dumps(evidence,default=str,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(evidence,default=str,ensure_ascii=False))


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path);parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--recalculate',action='store_true');args=parser.parse_args()
    run(args.source,args.output,args.recalculate)
