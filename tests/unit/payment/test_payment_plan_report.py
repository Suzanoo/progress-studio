from decimal import Decimal as D

import pytest

from progress_studio.domain.payment_plan_report import Leaf, Period, ReportSource, Row
from progress_studio.domain.payment_requirement_states import requirement_states


def source(a, b, amounts=(D(1000), D(50))):
    periods = tuple(Period(f'P{i+1:02}', None, None, ()) for i in range(len(a)))
    leaves = {}
    for aid, seq, amount in zip(('A', 'B'), (a, b), amounts):
        entries = {p.id: (D(v), 11, j+6) for j, (v, p) in enumerate(zip(seq, periods)) if v is not None}
        leaves[aid] = Leaf(amount, tuple(requirement_states([p.id for p in periods], entries)))
    rows = (Row('root', None, '', 'Project', 8, ('A','B'), False),
            Row('W', 'root', '10', 'Works', 9, ('A',), False),
            Row('A', 'W', '10.1', 'Pour', 10, ('A',), True),
            Row('B', 'root', '11', 'Finish', 11, ('B',), True))
    return ReportSource('p','demo','sha','monetary',periods,rows,leaves,())


def test_threshold_adjustment_reopening_and_recompletion():
    values = ['0.99989999','0.9999',None,'0.99995','0.99991','0.86','1',None]
    s = source(values,[None]*8)
    result = s.select(['W','A'], revision='sha')
    assert result.included_in == {'A':'W'}
    v = result.selected.values
    assert [x.transition for x in v] == [
        'open','completed','suppressed','adjusted_complete','adjusted_complete','reopened','recompleted','suppressed']
    assert v[0].display_fraction < 1
    assert v[1].display_fraction == 1 and v[1].money == D('999.9')
    assert v[1].delta == D('.00000001') and v[3].movement == D('.05')
    assert v[4].period_visible and v[4].cumulative_visible and v[4].correction
    assert not v[2].period_visible and not v[2].cumulative_visible
    assert result.retained_money[2] == D('999.9') and result.retained_money[-1] == 1000
    assert sum(x.movement for x in v) == v[-1].money


@pytest.mark.parametrize('amount,expected',[(D(50),None),(D(0),D(1000)),(None,None)])
def test_unknown_and_zero_weight(amount,expected):
    s = source(['1'],[None],(D(1000),amount))
    x = s.select(['root','W','A'], revision='sha')
    assert x.full_project and x.included_in == {'W':'root','A':'W'}
    v = x.project.values[0]
    assert v.money == expected and v.known_money == 1000
    assert v.known_leaves == 1 and v.known_basis == 1000
    if expected is None:
        assert v.transition == 'unresolved' and v.cumulative_visible


def test_before_first_explicit_zero_and_revision_binding():
    s = source([None,'0',None,'0.6'],['1',None,None,None],(D(0), D(30)))
    assert [v.origin for v in s.leaves['A'].schedule] == ['BeforeFirst','Explicit','Carried','Explicit']
    x = s.select(['A'],revision='sha')
    assert x.row_values['A'].values[-1].fraction == D('.6')
    assert x.selected.values[-1].money == 0 and x.selected.values[-1].fraction is None
    assert x.unselected.values[-1].money == x.project.values[-1].money == 30
    assert not x.full_project and not s.select([],revision='sha').selected.activities
    with pytest.raises(ValueError, match='revision'):
        s.select(['A'], revision='other')
    with pytest.raises(ValueError, match='Unknown'):
        s.select(['deleted'],revision='sha')
