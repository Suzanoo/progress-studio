"""Read-only real-source BP-1 oracle proof; writes evidence JSON, not a BOQ workbook."""
import argparse
from dataclasses import asdict
from decimal import Decimal as D
from pathlib import Path
import hashlib,json,sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from progress_studio.services.boq_payment_projection import BOQPaymentProjectionService


def run(session,payment,boq,output):
    source_hashes={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in (session,payment,boq)}
    result=BOQPaymentProjectionService().calculate(session,payment,accept_revision=True,boq_path=boq)
    # Independent oracle: raw session records + workbook cells. No production
    # projection, derived-state or catalog helper is used for expected numbers.
    raw=json.loads(session.read_text(encoding='utf-8'));pairs={}
    for a in raw['allocations']:pairs.setdefault(a['boq_key'],[]).append((a['activity_id'],D(str(a['share_percent']))/100))
    w=load_workbook(boq,read_only=True,data_only=True);brows=list(w[raw['boq_sheet']].values);h={v:i for i,v in enumerate(brows[0])}
    amounts={f"{raw['boq_sheet']}|{i}|{i}":D(str(r[h['Amount']])) for i,r in enumerate(brows[1:],2) if isinstance(r[h['Amount']],(int,float))};w.close()
    w=load_workbook(payment,read_only=True,data_only=True);s=w['Payment Input'];rows=list(s.values)
    columns=[(i,v) for i,v in enumerate(rows[5]) if isinstance(v,str) and v in result.periods]
    states={};entries=0
    for row in rows[7:]:
        if row[0]!='ACT':continue
        seq=[row[i] for i,_ in columns];entries+=sum(x is not None for x in seq)
        latest=D(0);values=[];known=any(x is not None for x in seq)
        for value in seq:
            if value is not None:latest=D(str(value))
            values.append(latest if known else None)
        states[row[2]]=values
    w.close();verified=0;full=unmapped=0;totals=[D(0)]*len(result.periods)
    for item in result.items:
        if item.item.role!='item':continue
        key=item.item.key;links=pairs.get(key,[])
        if not links:
            unmapped+=1;assert all(p.cumulative_amount is None for p in item.periods);continue
        assert sum(s for _,s in links)==1;full+=1;previous=D(0)
        for i,period in enumerate(item.periods):
            expected=sum(share*states[a][i] for a,share in links)
            assert abs(period.cumulative_fraction-expected)<D('1e-20')
            assert abs(period.cumulative_amount-amounts[key]*expected)<D('1e-15')
            assert abs(period.incremental_fraction-(expected-previous))<D('1e-20')
            totals[i]+=period.cumulative_amount;previous=expected;verified+=1
        assert abs(sum(x.incremental_amount for x in item.periods)-item.periods[-1].cumulative_amount)<D('1e-15')
    assert all(hashlib.sha256(p.read_bytes()).hexdigest()==source_hashes[str(p)] for p in (session,payment,boq))
    summary={'verified_item_periods':verified,'full_items':full,'unmapped_items':unmapped,'allocations':len(raw['allocations']),
        'activities':len(states),'requirements':entries,'periods':result.periods,'complete_mapped_amounts':totals,
        'diagnostics':[asdict(d) for d in result.diagnostics],'provenance':result.provenance,'sources_unchanged':True,
        'sample':asdict(next(x for x in result.items if x.item.key=='Project|5|5'))}
    output.parent.mkdir(parents=True,exist_ok=True);output.write_text(json.dumps(summary,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
    print(json.dumps({k:v for k,v in summary.items() if k not in ('sample','diagnostics','provenance')},default=str))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('session',type=Path);p.add_argument('payment',type=Path);p.add_argument('boq',type=Path);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    if a.output.resolve() in {x.resolve() for x in (a.session,a.payment,a.boq)}:p.error('Evidence output must not overwrite an input')
    run(a.session,a.payment,a.boq,a.output)
