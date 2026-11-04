from dataclasses import replace
from decimal import Decimal
from pathlib import Path
import xml.etree.ElementTree as ET

import pytest

from progress_studio.services.xml_writeback_service import WritebackRequest, XmlWritebackService
from progress_studio.infrastructure.schedule_xml.p6_amount_writer import P6AmountDocument
from progress_studio.infrastructure.excel.xml_writeback_reader import read_workbook_amounts
from tests.fixtures.xml_writeback import workbook, p6, edit_package, text_cell, P6_NS


@pytest.fixture
def setup(tmp_path):
    wb = workbook(tmp_path/'source.xlsx')
    xml = p6(tmp_path/'target.xml')
    request = WritebackRequest(wb,xml,'p6:udf:47',monetary_confirmed=True)
    return XmlWritebackService(), request


def change_xml(request, old, new):
    request.target_xml.write_bytes(request.target_xml.read_bytes().replace(old.encode(),new.encode()))


def assert_blocked(service, request, tmp_path, message):
    preview = service.assess(request)
    assert not preview.valid
    assert message.lower() in '\n'.join(preview.errors).lower(), preview.errors
    with pytest.raises(ValueError):
        service.export(preview,tmp_path/'blocked.xml',reviewed=True)
    assert not (tmp_path/'blocked.xml').exists()


def test_precise_subset_review_and_byte_preservation(setup,tmp_path):
    service,request=setup
    wb_before=request.workbook.read_bytes(); xml_before=request.target_xml.read_bytes()
    assessment=service.assess(request)
    assert assessment.valid,assessment.errors
    assert assessment.source_count==3
    assert assessment.replacement_count==2
    assert assessment.preserved_milestone_count==1
    assert assessment.target_only_ids==('NKC2-A1',)
    assert assessment.zero_count==1
    assert assessment.write_total==Decimal('12.34567890123456789')
    assert assessment.before_known_total==10
    output=tmp_path/'output.xml'
    service.export(assessment,output,reviewed=True)
    expected=xml_before.replace(b'<CostValue>1</CostValue>',b'<CostValue>12.34567890123456789</CostValue>').replace(b'<CostValue>9</CostValue>',b'<CostValue>0</CostValue>')
    assert output.read_bytes()==expected
    assert request.workbook.read_bytes()==wb_before
    assert request.target_xml.read_bytes()==xml_before
    repeated=service.assess(replace(request,target_xml=output))
    assert repeated.replacement_count==0
    assert repeated.unchanged_count==2
    second=tmp_path/'second.xml'
    service.export(repeated,second,reviewed=True)
    assert second.read_bytes()==output.read_bytes()


@pytest.mark.parametrize('raw',['-1','NaN','Infinity','1e309','1e-309','',None,'#VALUE!','1_000','1,000','1e99999'])
def test_rejects_invalid_amounts_even_when_other_rows_valid(setup,tmp_path,raw):
    service,request=setup
    workbook(request.workbook,amounts=(raw,'1','0'))
    assert_blocked(service,request,tmp_path,'Amount')


@pytest.mark.parametrize('kind',['e','b','str','inlineStr','s','d'])
def test_rejects_nonnumeric_cell_types(setup,tmp_path,kind):
    service,request=setup
    workbook(request.workbook,kinds=(kind,None,None))
    assert_blocked(service,request,tmp_path,'numeric')


def test_all_zero_subset_is_valid(setup,tmp_path):
    service,request=setup
    workbook(request.workbook,amounts=('0','0','0'))
    a=service.assess(request)
    assert a.valid and a.write_total==0 and a.zero_count==2
    service.export(a,tmp_path/'zeros.xml',reviewed=True)


def test_formula_requires_recalculation_confirmation_and_warns(setup,tmp_path):
    service,request=setup
    workbook(request.workbook,formulas=('1+2',None,None),amounts=('3','0','0'))
    assert_blocked(service,request,tmp_path,'recalculation')
    a=service.assess(replace(request,recalculated_and_saved=True))
    assert a.valid and a.formula_count==1 and a.write_total==3
    assert any('freshness cannot' in w for w in a.warnings)
    service.export(a,tmp_path/'formula.xml',reviewed=True)


@pytest.mark.parametrize('value,kind',[(None,None),('#REF!','e'),('NaN',None),('3','str')])
def test_invalid_formula_cache_never_accepted(setup,tmp_path,value,kind):
    service,request=setup
    workbook(request.workbook,formulas=('1+2',None,None),amounts=(value,'0','0'),kinds=(kind,None,None))
    assert_blocked(service,replace(request,recalculated_and_saved=True),tmp_path,'Amount')


@pytest.mark.parametrize('ids,message',[(('A1','A1','M1'),'Duplicate'),(('A1','a1','M1'),'case-colliding'),(('', 'A2','M1'),'Blank'),(('UNKNOWN','A2','M1'),'No exact'),(('a1','A2','M1'),'No exact')])
def test_bad_source_identities(setup,tmp_path,ids,message):
    service,request=setup
    workbook(request.workbook,ids=ids)
    assert_blocked(service,request,tmp_path,message)


@pytest.mark.parametrize('old,new,message',[
    ('<Id>A2</Id>','<Id>A1</Id>','duplicate'),
    ('<Id>A2</Id>','<Id>a1</Id>','case-colliding'),
    ('<Id>A2</Id>','<Id/>','Blank'),
    ('<ObjectId>11</ObjectId>','<ObjectId>10</ObjectId>','ObjectId'),
    ('<Id>A1</Id>','<Id>A1</Id><Id>A1</Id>','duplicate'),
    ('<Type>Task Dependent</Type>','<Type>Unknown</Type>','Type'),
    ('<Type>Resource Dependent</Type>','<Type>Finish Milestone</Type>','Conflicting'),
    ('</APIBusinessObjects>','<Project/></APIBusinessObjects>','one direct Project'),
    ('<DataType>Cost</DataType>','<DataType>Double</DataType>','Cost UDF'),
    ('<SubjectArea>Activity</SubjectArea>','<SubjectArea>Project</SubjectArea>','Cost UDF'),
    ('<CostValue>1</CostValue>','<DoubleValue>1</DoubleValue>','incompatible'),
    ('<CostValue>1</CostValue>','<CostValue>1</CostValue><CostValue>2</CostValue>','duplicate'),
    ('<CostValue>1</CostValue>','<CostValue xsi:nil="true">1</CostValue>','nil'),
    ('<CostValue>1</CostValue>','<CostValue>-5</CostValue>','negative'),
    ('<UDF><TypeObjectId>47</TypeObjectId><CostValue>1</CostValue></UDF>',
     '<UDF><TypeObjectId>47</TypeObjectId><CostValue>1</CostValue></UDF><UDF><TypeObjectId>47</TypeObjectId><CostValue>2</CostValue></UDF>','duplicate'),
])
def test_invalid_target_is_atomic(setup,tmp_path,old,new,message):
    service,request=setup
    change_xml(request,old,new)
    assert_blocked(service,request,tmp_path,message)


def test_duplicate_udf_declaration(setup,tmp_path):
    service,request=setup
    text=request.target_xml.read_text(encoding="utf-8")
    declaration=text[text.index('<UDFType>'):text.index('</UDFType>')+10]
    change_xml(request,'</UDFType>','</UDFType>'+declaration)
    assert_blocked(service,request,tmp_path,'duplicate')


@pytest.mark.parametrize('old,new',[
    ('<CostValue>1</CostValue>',''),
    ('<CostValue>1</CostValue>','<CostValue/>'),
    ('<CostValue>1</CostValue>','<CostValue xsi:nil="true"/>'),
    ('<CostValue>1</CostValue>','<v:CostValue xmlns:v="'+P6_NS+'" xsi:nil="true"/>'),
    ('<UDF><TypeObjectId>47</TypeObjectId><CostValue>1</CostValue></UDF>',''),
])
def test_missing_and_nil_values_are_reviewed_replacements(setup,tmp_path,old,new):
    service,request=setup
    change_xml(request,old,new)
    a=service.assess(request)
    assert a.valid,a.errors
    assert a.before_missing_count==1
    output=tmp_path/'out.xml'
    service.export(a,output,reviewed=True)
    doc=P6AmountDocument(output.read_bytes())
    assert doc.value_node('A1','p6:udf:47')[2]=='12.34567890123456789'


def test_new_udf_avoids_all_object_and_dangling_reference_ids(setup,tmp_path):
    service,request=setup
    p6(request.target_xml,declared=False)
    change_xml(request,'</Project>','<x:External><TypeObjectId>900</TypeObjectId></x:External></Project>')
    request=replace(request,field_identity=None,create_progress_studio_amount=True)
    a=service.assess(request)
    assert a.valid,a.errors
    assert a.field.identity=='p6:udf:901' and a.creates_field
    output=tmp_path/'new.xml';service.export(a,output,reviewed=True)
    doc=P6AmountDocument(output.read_bytes())
    assert doc.value_node('M1',a.field.identity)[2] is None
    assert doc.value_node('NKC2-A1',a.field.identity)[2] is None
    assert len(doc.definitions)==1
    assert_blocked(service,replace(request,target_xml=output),tmp_path,'already exists')


def test_creation_does_not_reuse_conflicting_named_text_field(setup,tmp_path):
    service,request=setup
    change_xml(request,'<Title>Budget</Title>','<Title>Progress Studio Amount</Title>')
    change_xml(request,'<DataType>Cost</DataType>','<DataType>Text</DataType>')
    assert_blocked(service,replace(request,field_identity=None,create_progress_studio_amount=True),tmp_path,'already exists')


@pytest.mark.parametrize('prefix,shifted',[(True,False),(False,True),(True,True)])
def test_prefixes_utf8_and_shifted_columns(setup,tmp_path,prefix,shifted):
    service,request=setup
    workbook(request.workbook,shifted=shifted)
    p6(request.target_xml,prefix=prefix)
    change_xml(request,'Combined project','โครงการ ทดสอบ')
    a=service.assess(request);assert a.valid,a.errors
    service.export(a,tmp_path/'out.xml',reviewed=True)
    assert 'โครงการ ทดสอบ' in (tmp_path/'out.xml').read_text(encoding="utf-8")


def test_unknown_source_classification_blocks_without_guessing(setup,tmp_path):
    service,request=setup
    workbook(request.workbook,flags=None)
    assert_blocked(service,request,tmp_path,'Unknown source milestone')


def test_milestone_money_is_never_replaced_by_nonzero_main(setup,tmp_path):
    service,request=setup
    workbook(request.workbook,amounts=('5','0','700'))
    a=service.assess(request);assert a.valid
    assert a.source_total==705 and a.write_total==5
    service.export(a,tmp_path/'out.xml',reviewed=True)
    assert b'<CostValue>500.50</CostValue>' in (tmp_path/'out.xml').read_bytes()


@pytest.mark.parametrize('transform,message',[
    (lambda s:s.replace('Amount</t>','Other</t>'),'headers'),
    (lambda s:s.replace('</row>',text_cell('Z4','Amount')+'</row>',1),'Duplicate required'),
    (lambda s:s.replace('<t>P</t>','<t>?</t>',1),'P or A'),
    (lambda s:s.replace('<t>Activity</t>','<t>Unknown</t>',1),'non-Activity'),
    (lambda s:s.replace('<c r="C5" t="inlineStr">','<c r="C5" t="inlineStr"><f>"A1"</f>',1),'literal'),
])
def test_headers_and_identity_evidence(setup,tmp_path,transform,message):
    service,request=setup
    edit_package(request.workbook,'xl/worksheets/sheet1.xml',transform)
    assert_blocked(service,request,tmp_path,message)


def test_corrupt_packages_and_xml_return_assessment_errors(setup,tmp_path):
    service,request=setup
    request.workbook.write_bytes(b'not a workbook')
    assert_blocked(service,request,tmp_path,'Workbook structure')
    workbook(request.workbook)
    request.target_xml.write_bytes(b'<broken>')
    assert_blocked(service,request,tmp_path,'Invalid P6 XML')


def test_dtd_and_non_utf8_rejected(setup,tmp_path):
    service,request=setup
    change_xml(request,'<APIBusinessObjects','<!DOCTYPE APIBusinessObjects [<!ENTITY money "900">]><APIBusinessObjects')
    assert_blocked(service,request,tmp_path,'DTD')
    p6(request.target_xml)
    change_xml(request,'encoding="utf-8"','encoding="iso-8859-1"')
    assert_blocked(service,request,tmp_path,'UTF-8')


def test_explicit_monetary_and_field_selection_required(setup,tmp_path):
    service,request=setup
    assert_blocked(service,replace(request,monetary_confirmed=False),tmp_path,'monetary')
    assert_blocked(service,replace(request,field_identity=None),tmp_path,'Explicitly select')
    assert_blocked(service,replace(request,create_progress_studio_amount=True),tmp_path,'not both')


def test_precision_sum_exceeds_default_decimal_context(setup):
    service,request=setup
    workbook(request.workbook,amounts=('123456789012345678901234567890.123456789','0.000000001','0'))
    a=service.assess(request)
    assert a.valid,a.errors
    assert str(a.write_total)=='123456789012345678901234567890.123456790'


def test_info_ev_classification_conflict_is_not_ignored(setup,tmp_path):
    from tests.fixtures.xml_writeback import sheet, number_cell, write_package
    from zipfile import ZipFile
    service,request=setup
    with ZipFile(request.workbook) as z:
        main=z.read('xl/worksheets/sheet1.xml').decode()
        ev=z.read('xl/worksheets/sheet2.xml').decode()
    info=sheet([(1,[text_cell('A1','Calendar Duration Milestone: A1'),number_cell('B1','1',kind='b')])])
    write_package(request.workbook,{'main':main,'EV Monetary Inputs':ev,'Info':info})
    assert_blocked(service,request,tmp_path,'conflicting Info/EV')


def test_duplicate_main_cell_cannot_hide_invalid_amount(setup,tmp_path):
    service,request=setup
    edit_package(request.workbook,'xl/worksheets/sheet1.xml',
                 lambda s:s.replace('</row>','<c r="D4"><v>1</v></c></row>',1))
    assert_blocked(service,request,tmp_path,'Duplicate or invalid worksheet cell')


def test_milestone_conflict_and_invalid_milestone_amount_block(setup,tmp_path):
    service,request=setup
    workbook(request.workbook,flags=(False,False,False))
    assert_blocked(service,request,tmp_path,'Conflicting')
    workbook(request.workbook,amounts=('1','2',None))
    assert_blocked(service,request,tmp_path,'Amount')


def test_duplicate_sheet_relationship_is_ambiguous(setup,tmp_path):
    service,request=setup
    edit_package(request.workbook,'xl/_rels/workbook.xml.rels',lambda s:s.replace('Id="rId2"','Id="rId1"'))
    assert_blocked(service,request,tmp_path,'relationship identities')


def test_cell_address_must_belong_to_its_row(setup,tmp_path):
    service,request=setup
    edit_package(request.workbook,'xl/worksheets/sheet1.xml',lambda s:s.replace('r="D5"','r="D6"'))
    assert_blocked(service,request,tmp_path,'cell/row address')
