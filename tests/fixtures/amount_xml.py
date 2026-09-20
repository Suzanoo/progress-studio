"""Small source-accurate custom field inputs with identities unlike customer samples."""
import xml.etree.ElementTree as ET


def amount_xml(tmp_path, source, values=('10.1234567890123', '20.9876543210987', '0', '999')):
    p6 = source == 'p6'
    namespace = ('http://xmlns.oracle.com/Primavera/P6Professional/V24.12/API/BusinessObjects'
                 if p6 else 'http://schemas.microsoft.com/project')
    ns = '{' + namespace + '}'
    root = ET.Element(ns + ('APIBusinessObjects' if p6 else 'Project'))
    project = ET.SubElement(root, ns + 'Project') if p6 else root
    def fields(parent, **values):
        for key, value in values.items():
            ET.SubElement(parent, ns + key).text = value
    fields(project, Name='Amount fixture')
    if p6:
        for identity, title, dtype in [('91', 'Contract allocation', 'Double'), ('92', 'Other value', 'Integer'), ('93', 'Amount text', 'Text')]:
            node = ET.SubElement(root, ns + 'UDFType')
            fields(node, ObjectId=identity, Title=title, DataType=dtype, SubjectArea='Activity')
        wbs = ET.SubElement(project, ns + 'WBS')
        fields(wbs, ObjectId='1', Code='1', Name='Work', SequenceNumber='1')
    else:
        definitions = ET.SubElement(root, ns + 'ExtendedAttributes')
        for identity, name, alias, dtype in [('91', 'Number17', 'Contract allocation', '5'), ('92', 'Number26', 'Other value', '5'), ('93', 'Text14', 'Amount text', '7')]:
            node = ET.SubElement(definitions, ns + 'ExtendedAttribute')
            fields(node, FieldID=identity, FieldName=name, Alias=alias, CFType=dtype)
        tasks = ET.SubElement(root, ns + 'Tasks')
        summary = ET.SubElement(tasks, ns + 'Task')
        fields(summary, ID='1', UID='1', Name='Work', WBS='1', OutlineLevel='1', Summary='1')
    for index, value in enumerate(values):
        milestone = index == len(values) - 1
        aid = f'A{index+1}'
        name = 'Handover' if milestone else f'Work {index+1}'
        start = '2026-01-01T08:00:00'; finish = '2026-01-23T17:00:00'
        if p6:
            activity = ET.SubElement(project, ns + 'Activity')
            fields(activity, Id=aid, Name=name, WBSObjectId='1', PlannedStartDate=finish if milestone else start,
                   PlannedFinishDate=finish, PlannedDuration='0' if milestone else str(8*(index+1)),
                   Type='Finish Milestone' if milestone else 'Task Dependent', PercentComplete='0')
            for identity, raw, tag in [('91', value, 'DoubleValue'), ('92', '80', 'IntegerValue'), ('93', '1000', 'TextValue')]:
                if raw is not None:
                    node = ET.SubElement(activity, ns + 'UDF')
                    fields(node, TypeObjectId=identity, **{tag: raw})
        else:
            activity = ET.SubElement(tasks, ns + 'Task')
            fields(activity, ID=str(index+2), UID=str(index+2), Name=name, WBS='1', OutlineLevel='2', Summary='0',
                   Start=finish if milestone else start, Finish=finish, Duration='PT0H' if milestone else f'PT{8*(index+1)}H',
                   Milestone='1' if milestone else '0', PercentComplete='0', Cost='999999')
            for identity, raw in [('188743731', aid), ('91', value), ('92', '80'), ('93', '1000')]:
                if raw is not None:
                    node = ET.SubElement(activity, ns + 'ExtendedAttribute')
                    fields(node, FieldID=identity, Value=raw)
    path = tmp_path / f'{source}.xml'
    ET.ElementTree(root).write(path, encoding='utf-8', xml_declaration=True)
    return path
