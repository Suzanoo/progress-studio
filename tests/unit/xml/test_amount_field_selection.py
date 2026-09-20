from decimal import Decimal
from dataclasses import replace
import pytest

from progress_studio.infrastructure.schedule_xml import NormalizedScheduleXmlReader
from progress_studio.services.xml_amount_service import preview_amounts, assign_xml_amounts
from progress_studio.services.weighting import assign_dummy_weights
from tests.fixtures.amount_xml import amount_xml


@pytest.mark.parametrize('source', ['p6', 'msp'])
def test_discovery_is_typed_not_name_or_position_and_requires_selection(tmp_path, source):
    _, rows, fields = NormalizedScheduleXmlReader().read_with_amount_fields(amount_xml(tmp_path, source))
    assert len(fields) == 2
    assert {f.name for f in fields} == {'Contract allocation', 'Other value'}
    assert all(r.amount is None for r in rows)
    with pytest.raises(ValueError, match='Select an XML Amount field'):
        preview_amounts(rows, fields, None)
    with pytest.raises(ValueError, match='missing, not numeric'):
        preview_amounts(rows, fields, 'unknown')
    preview = assign_xml_amounts(rows, fields, fields[0].identity)
    assert preview.valid
    assert (preview.ordinary_count, preview.positive_count, preview.zero_count, preview.milestone_count) == (3, 2, 1, 1)
    assert preview.total == Decimal('31.1111111101110')
    activities = [r for r in rows if not r.is_summary]
    assert [r.activity_id for r in activities] == ['A1','A2','A3','A4']
    assert [r.amount for r in activities] == [float('10.1234567890123'),float('20.9876543210987'),0,0]
    second = assign_xml_amounts(rows, fields, fields[1].identity)
    assert second.total == 240
    # Even duplicate display titles do not select the wrong field.
    duplicate_names = tuple(replace(f, name='Same title') for f in fields)
    assert preview_amounts(rows, duplicate_names, fields[1].identity).total == 240


@pytest.mark.parametrize('source', ['p6','msp'])
@pytest.mark.parametrize('raw,reason', [(None,'missing'),('','missing'),('bad','non-numeric'),('-0.01','negative'),('NaN','non-finite'),('Infinity','non-finite'),('1e309','numeric range'),('1e-999','numeric range'),('1_000','non-numeric')])
def test_bad_selected_values_are_actionable_and_assignment_is_atomic(tmp_path, source, raw, reason):
    _, rows, fields = NormalizedScheduleXmlReader().read_with_amount_fields(amount_xml(tmp_path, source, (raw,'20','0','0')))
    preview = preview_amounts(rows, fields, fields[0].identity)
    assert not preview.valid
    assert 'A1 (Work 1)' in preview.errors[0]
    assert 'Contract allocation' in preview.errors[0]
    assert reason in preview.errors[0]
    with pytest.raises(ValueError, match=reason):
        assign_xml_amounts(rows, fields, fields[0].identity)
    assert all(r.amount is None for r in rows)
    # An unrelated invalid Amount field must not affect accepted dummy modes.
    assign_dummy_weights(rows,'equal')
    assert [r.amount for r in rows if not r.is_summary] == [1,1,1,0]
    assign_dummy_weights(rows,'duration')
    assert [r.amount for r in rows if not r.is_summary] == [8,16,24,0]


@pytest.mark.parametrize('source', ['p6','msp'])
def test_total_and_duplicate_values_are_rejected(tmp_path, source):
    _, rows, fields = NormalizedScheduleXmlReader().read_with_amount_fields(amount_xml(tmp_path,source,('0','0','0','999')))
    assert 'total' in preview_amounts(rows,fields,fields[0].identity).errors[0]
    activity = next(r for r in rows if not r.is_summary)
    activity.amount_field_values += ((fields[0].identity,'10'),)
    assert 'duplicate' in preview_amounts(rows,fields,fields[0].identity).errors[0]
    with pytest.raises(ValueError,match='ambiguous'):
        preview_amounts(rows,(*fields,fields[0]),fields[0].identity)


@pytest.mark.parametrize('source', ['p6','msp'])
def test_milestone_field_is_not_required_and_positive_values_are_not_rounded(tmp_path,source):
    _,rows,fields=NormalizedScheduleXmlReader().read_with_amount_fields(amount_xml(tmp_path,source,('0.00001','0.00002','0',None)))
    p=assign_xml_amounts(rows,fields,fields[0].identity)
    assert p.total == Decimal('0.00003')
    assert p.rows[-1].value == 0
    assert next(r for r in rows if r.activity_id=='A1').amount == 0.00001
