from datetime import date
from hashlib import sha256

import pytest
from openpyxl import load_workbook
from openpyxl.comments import Comment
from progress_studio.infrastructure.excel.payment_plan_source import read_payment_plan_source
from tests.integration.payment.test_payment_position_engine import _progress_workbook


@pytest.fixture
def book(tmp_path):
    path = _progress_workbook(tmp_path/'source.xlsx')
    wb = load_workbook(path)
    main = wb['main']
    main['H5'],main['H6'],main['H8'] = 0,1,1
    main['O6'],main['O8'] = 100,50
    info = wb.create_sheet('Info')
    for row in [('Project','Demo'),('Creation Weight Basis','amount'),
                ('Creation Amount Field identity','cost'),('Amount Initialization Source','xml_field')]:
        info.append(row)
    ws = wb.create_sheet('Payment Input')
    for i, val in enumerate(('Type','WBS','Activity ID','Activity Name','Amount','P01','P03'),1):
        ws.cell(6,i,val)
    for i,row in enumerate([('WBS',None,None,'Demo'),('ACT',None,'A1000','Foundation',100,0,1),
                            ('ACT',None,'A1010','Structure',50,None,'25%')],8):
        for j,v in enumerate(row,1):
            ws.cell(i,j,v)
    wb.create_sheet('Payment')['F4'].comment = Comment('P01 Planned Eligible Date: 1990-01-01','test')
    wb.save(path)
    wb.close()
    return path


def edit(path,sheet,addr,value):
    wb = load_workbook(path)
    wb[sheet][addr] = value
    wb.save(path)
    wb.close()


def test_project_wide_dates_ignore_stale_marker_and_track_edits(book):
    before = sha256(book.read_bytes()).hexdigest()
    s = read_payment_plan_source(book, expected_revision=before)
    assert s.periods[0].eligible == date(2026,3,2)
    assert s.periods[1].eligible == date(2026,3,30)
    assert s.select(['ACT/A1010'],revision=before).selected.values[1].money == 12.5
    assert sha256(book.read_bytes()).hexdigest() == before
    edit(book,'Payment Input','F9',0.6)
    assert read_payment_plan_source(book).periods[0].eligible == date(2026,3,23)
    with pytest.raises(ValueError,match='revision'):
        read_payment_plan_source(book,expected_revision=before)


@pytest.mark.parametrize('sheet,addr,value,message',[
    ('Payment Input','E6','Old','Missing'),('Payment Input','E9',99,'Stale'),
    ('main','O6','=100','Formula'),('Payment Input','F9','=0','Formula'),
    ('Payment Input','F9',-0.2,'between'),('Payment Input','F9',True,'Invalid'),
    ('Payment Input','C10','A1000','Duplicate'),('Payment Input','D9','Changed','Hierarchy'),
    ('Info','B2','duration','monetary'),('Info','B4','manual_placeholder','monetary'),
    ('Payment Input','G6','P1','order'),('Payment Input','G6','P01','Duplicate')])
def test_source_diagnostics(book,sheet,addr,value,message):
    edit(book,sheet,addr,value)
    with pytest.raises(ValueError,match=message):
        read_payment_plan_source(book)


def test_unverified_plan_never_publishes_partial_date(book):
    edit(book,'main','R6','=0.10')
    s=read_payment_plan_source(book)
    assert all(p.eligible is None and p.issues for p in s.periods)


def test_missing_money_keeps_known_subtotal(book):
    edit(book,'main','O6',None)
    edit(book,'Payment Input','E9','Missing Amount')
    s=read_payment_plan_source(book)
    assert 'AmountUnavailable' in s.diagnostics[0]
    v=s.select(['PROJECT'],revision=s.revision).project.values[-1]
    assert v.money is None and v.known_money == 12.5 and v.fraction is None


def test_date_reversal_preserves_signed_adjustment(book):
    edit(book,'Payment Input','F9',1)
    edit(book,'Payment Input','G9',0)
    s=read_payment_plan_source(book)
    assert s.periods[1].eligible is None and 'precedes' in s.periods[1].issues[0]
    v=s.select(['ACT/A1000'],revision=s.revision).selected.values[1]
    assert v.correction and v.movement == -100


def test_empty_period_has_unresolved_date_but_carry_survives(book):
    edit(book,'Payment Input','G9',None)
    edit(book,'Payment Input','G10',None)
    s=read_payment_plan_source(book)
    assert s.periods[-1].eligible is None and s.periods[-1].issues
    assert s.leaves['A1000'].schedule[-1].origin=='Carried'


def test_stale_outline_detected(book):
    wb=load_workbook(book)
    wb['Payment Input'].row_dimensions[9].outlineLevel=2
    wb.save(book)
    wb.close()
    with pytest.raises(ValueError,match='outline'):
        read_payment_plan_source(book)


def test_reordered_headers_and_mapping_provenance(book):
    wb=load_workbook(book)
    wb.create_sheet('BOQ Activity Mapping')
    wb['Info']['B2']='equal'
    ws=wb['Payment Input']
    for row in (6,8,9,10):
        ws.cell(row,3).value,ws.cell(row,4).value=ws.cell(row,4).value,ws.cell(row,3).value
    wb.save(book)
    wb.close()
    s=read_payment_plan_source(book)
    assert s.monetary_origin=='BOQ Activity Mapping' and len(s.leaves)==2


def test_legacy_and_derived_workbooks_fail_with_recovery_step(book):
    wb=load_workbook(book)
    del wb['main']
    wb.save(book)
    wb.close()
    with pytest.raises(ValueError,match='embedded Payment source'):
        read_payment_plan_source(book)


def test_long_literal_percent_retains_threshold_side(book):
    from decimal import Decimal
    edit(book,'Payment Input','G9','99.98999999999999999999999999999999999%')
    s=read_payment_plan_source(book)
    v=s.select(['ACT/A1000'],revision=s.revision).selected.values[1]
    assert v.fraction==Decimal('0.9998999999999999999999999999999999999')
    assert v.transition=='open'
