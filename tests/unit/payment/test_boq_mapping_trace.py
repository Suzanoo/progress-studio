from dataclasses import replace
from decimal import Decimal
import pytest
from progress_studio.domain.boq_mapping_trace import mapping_trace
from progress_studio.domain.boq_manual_requirements import compose, ManualInput
from tests.unit.payment.test_boq_payment_projection import run


def test_relationships_are_period_independent_and_names_disambiguated():
    p=replace(run(),activity_names=(('A','Same name'),('B','Same name'),('C','No mapping')))
    t=mapping_trace(p)
    assert len(t.relationships)==2
    assert t.activities==(('A','Same name'),('B','Same name'),('C','No mapping'))
    assert dict(t.by_activity)['C']==()
    for key,indexes in t.by_boq:
        for i in indexes:
            edge=t.relationships[i]
            assert edge.boq_key==key and i in dict(t.by_activity)[edge.activity_id]
    assert t.capacity==2


def test_manual_composition_preserves_names_and_creates_no_mapping():
    p=replace(run(shares=()),activity_names=(('A','Actual name'),))
    result=compose(p,(ManualInput('b',(Decimal(0),)*3,9),))
    assert result.activity_names==p.activity_names
    assert result.mapped_projection is p
    assert mapping_trace(result).relationships==()


def test_contradictory_period_relationships_rejected():
    p=run();entry=p.items[0]
    changed=replace(entry,periods=(entry.periods[0],replace(entry.periods[1],contributions=()),entry.periods[2]))
    with pytest.raises(ValueError,match='differ'):mapping_trace(replace(p,items=(changed,)))
