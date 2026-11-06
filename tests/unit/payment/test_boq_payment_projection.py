from dataclasses import replace
from decimal import Decimal as D
from pathlib import Path
import pytest
from progress_studio.domain.boq_payment_projection import CatalogRow, project
from progress_studio.domain.mapping_models import AllocationRecord
from progress_studio.domain.payment_models import PaymentInputData, PaymentPeriodRequirements, PaymentRequirement


def item(key='b',amount=D(100),role='item'):
    return CatalogRow(key,key,'Project',2,2,'1.1',('a','b','c'),'Item','m2',1,amount,amount,'Valid',role)


def payment(a=(.2,None,.8),b=(None,.5,1)):
    periods=[]
    for i,p in enumerate(('P01','P03','P09')):
        req=tuple(PaymentRequirement(name,v[i],8+j,6+i) for j,(name,v) in enumerate((('A',a),('B',b))) if v[i] is not None)
        periods.append(PaymentPeriodRequirements(p,6+i,None,req))
    return PaymentInputData(Path('source.xlsx'),'Payment Input',tuple(periods),('A','B'),sum(len(p.requirements) for p in periods))


def run(a=(.2,None,.8),b=(None,.5,1),shares=(60,40),amount=D(100)):
    return project((item(amount=amount),),tuple(AllocationRecord('b',x,s) for x,s in zip(('A','B'),shares)),payment(a,b),('A','B'))


def test_split_conservation_and_trace():
    result=run();p=result.items[0].periods
    assert [r.cumulative_fraction for r in p]==list(map(D,('.12','.32','.88')))
    assert [r.incremental_fraction for r in p]==list(map(D,('.12','.20','.56')))
    assert sum(r.incremental_amount for r in p)==D(88)
    assert p[0].contributions[1].state=='BeforeFirst'
    assert p[1].contributions[0].state=='Carried'
    assert p[1].contributions[0].source_column==6
    assert result.periods==('P01','P03','P09')


def test_partial_unknown_and_unmapped_are_distinct():
    partial=run(shares=(60,)).items[0].periods[-1]
    assert partial.status=='PartialMapping' and partial.known_fraction==D('.48')
    assert partial.cumulative_fraction is None and partial.incremental_amount is None
    unknown=run(b=(None,None,None)).items[0].periods[-1]
    assert unknown.status=='IncompleteRequirement' and unknown.input_coverage==D('.6')
    assert unknown.cumulative_fraction is None
    empty=run(shares=()).items[0].periods[-1]
    assert empty.status=='Unmapped' and empty.known_fraction is None


def test_zero_blank_decreases_and_milestone_independence():
    r=run(a=(1,0,None),b=(1,1,1));p=r.items[0].periods
    assert [x.cumulative_fraction for x in p]==[D(1),D('.4'),D('.4')]
    assert p[1].incremental_fraction==D('-.6') and p[2].warnings==('DecreasingRequirement',)
    assert p[1].contributions[0].state=='Explicit' and p[2].contributions[0].state=='Carried'
    assert run(a=(None,None,None),b=(None,None,None)).items[0].periods[0].known_fraction is None
    # Activity Amount does not exist in this contract, so zero-weight milestones
    # with positive BOQ shares cannot accidentally disappear.
    assert run(a=(0,None,1),b=(0,None,1)).items[0].periods[-1].cumulative_fraction==1


@pytest.mark.parametrize('amount',[None,D(-10),D(0)])
def test_amount_does_not_invent_or_block_valid_percentages(amount):
    p=run(amount=amount).items[0].periods[-1]
    assert p.cumulative_fraction==D('.88')
    assert p.cumulative_amount==(D(0) if amount==0 else None)


@pytest.mark.parametrize('share',[0,-1,101,float('nan'),float('inf'),True])
def test_invalid_shares_fail(share):
    with pytest.raises(ValueError):run(shares=(share,))


@pytest.mark.parametrize('value',[-.1,1.1,float('nan'),float('inf'),True])
def test_invalid_requirements_fail(value):
    with pytest.raises(ValueError):run(a=(value,None,1))


def test_duplicate_unknown_and_overallocation():
    pay=payment();a=AllocationRecord('b','A',60)
    for allocations in [(a,a),(a,AllocationRecord('b','B',50)),(AllocationRecord('missing','A',10),),(AllocationRecord('b','missing',10),)]:
        with pytest.raises(ValueError):project((item(),),allocations,pay,('A','B'))
    with pytest.raises(ValueError):project((item(),item()),(),pay,('A','B'))
    with pytest.raises(ValueError):project((item(),),(),replace(pay,periods=(pay.periods[0],pay.periods[0])),('A','B'))


def test_rounding_and_leaf_only_catalog():
    rows=(item('group',D(999),'group'),item('b',D('0.01')),item('c',D('0.02')))
    r=project(rows,(AllocationRecord('b','A',100),AllocationRecord('c','A',100)),payment(a=(.333333,None,1)),('A','B'))
    assert r.items[0].periods==()
    assert sum(x.periods[-1].cumulative_amount for x in r.items if x.periods)==D('.03')
    for x in r.items[1:]:assert sum(p.incremental_amount for p in x.periods)==x.item.amount
    assert r.items[1].periods[0].cumulative_amount==D('.00333333')
