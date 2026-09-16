"""Publish only FF-2 owned parts into the original package.

Finance has no drawing/chart parts. Unlike restoring raw drawing numbers after a
whole-workbook round-trip, this narrow merge never reserializes existing owners.
The existing atomic-write and table-validation helpers remain the outer boundary.
"""
from copy import deepcopy
from io import BytesIO
import re
from pathlib import PurePosixPath
import posixpath
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED
from progress_studio.infrastructure.excel.finance_input_workbook import INPUT, DATA, VIEW

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/package/2006/relationships'
D='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C='http://schemas.openxmlformats.org/package/2006/content-types'


_NAMESPACES = {}

def _xml(data):
    root=ET.fromstring(data)
    _NAMESPACES[id(root)]=dict(value for _,value in ET.iterparse(BytesIO(data),events=['start-ns']))
    return root

def _bytes(root):
    namespaces=_NAMESPACES.get(id(root),{})
    for prefix,uri in namespaces.items():
        if not re.fullmatch(r'ns[0-9]+',prefix): ET.register_namespace(prefix,uri)
    data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    # Preserve declarations referenced only in markup-compatibility attribute
    # values. ElementTree otherwise drops unused namespace declarations.
    end=data.index(b'>',data.index(b'?>')+2)
    for prefix,uri in namespaces.items():
        declaration=('xmlns'+(':'+prefix if prefix else '')+'=').encode()
        if declaration not in data[:end]:
            data=data[:end]+b' '+declaration+b'"'+uri.encode()+b'"'+data[end:]
            end=data.index(b'>',data.index(b'?>')+2)
    return data
def _resolve(owner,target):return posixpath.normpath(posixpath.join(posixpath.dirname(owner),target)) if not target.startswith('/') else target[1:]
def _relpath(path):
    p=PurePosixPath(path);return str(p.parent/'_rels'/(p.name+'.rels'))


def _sheets(parts):
    rels={r.get('Id'):_resolve('xl/workbook.xml',r.get('Target')) for r in _xml(parts['xl/_rels/workbook.xml.rels'])}
    return {s.get('name'):(s,rels[s.get('{'+D+'}id')]) for s in _xml(parts['xl/workbook.xml']).find('{'+S+'}sheets')}


def _merge_styles(original,generated):
    old=_xml(original);new=_xml(generated)
    def group(root,name):
        node=root.find('{'+S+'}'+name)
        if node is None:
            node=ET.Element('{'+S+'}'+name)
            # OOXML order, including optional collections.
            order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
            index=next((i for i,n in enumerate(root) if n.tag.split('}')[-1] in order and order.index(n.tag.split('}')[-1])>order.index(name)),len(root))
            root.insert(index,node)
        return node
    def merge(name,transform=lambda n:n):
        dest=group(old,name);mapping={}
        for i,node in enumerate(group(new,name)):
            node=transform(deepcopy(node));key=ET.tostring(node)
            found=next((j for j,n in enumerate(dest) if ET.tostring(n)==key),None)
            if found is None:found=len(dest);dest.append(node)
            mapping[i]=found
        dest.set('count',str(len(dest)));return mapping
    formats={}
    dest=group(old,'numFmts')
    for node in group(new,'numFmts'):
        existing=next((n for n in dest if n.get('formatCode')==node.get('formatCode')),None)
        if existing is None:
            existing=deepcopy(node);existing.set('numFmtId',str(max([163]+[int(n.get('numFmtId')) for n in dest])+1));dest.append(existing)
        formats[int(node.get('numFmtId'))]=int(existing.get('numFmtId'))
    dest.set('count',str(len(dest)))
    maps={k:merge(v) for k,v in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]}
    def xf(node):
        for key,mapping in maps.items():
            if key in node.attrib:node.set(key,str(mapping[int(node.get(key))]))
        fmt=int(node.get('numFmtId','0'));node.set('numFmtId',str(formats.get(fmt,fmt)))
        return node
    xfs=merge('cellStyleXfs',xf)
    def cellxf(node):
        node=xf(node)
        if 'xfId' in node.attrib:node.set('xfId',str(xfs[int(node.get('xfId'))]))
        return node
    styles=merge('cellXfs',cellxf)
    return _bytes(old),styles


def merge_finance_parts(source,rendered):
    """Mutate temporary rendered file only, preserving every unowned source byte."""
    with ZipFile(source) as z: original={n:z.read(n) for n in z.namelist()}
    with ZipFile(rendered) as z: generated={n:z.read(n) for n in z.namelist()}
    result=dict(original)
    result['xl/styles.xml'],style_map=_merge_styles(original['xl/styles.xml'],generated['xl/styles.xml'])
    old_sheets=_sheets(original);new_sheets=_sheets(generated)
    book=_xml(original['xl/workbook.xml']); sheets=book.find('{'+S+'}sheets')
    rels=_xml(original['xl/_rels/workbook.xml.rels']);types=_xml(original['[Content_Types].xml'])
    original_types={n.get('PartName'):n for n in types}
    generated_types={n.get('PartName'):n for n in _xml(generated['[Content_Types].xml'])}
    table_ids=[int(_xml(v).get('id','0')) for k,v in original.items() if k.startswith('xl/tables/') and k.endswith('.xml') and '/_rels/' not in k]
    next_table_id=max(table_ids,default=0)+1
    used={r.get('Id') for r in rels}
    def add_type(path,kind):
        key='/'+path
        if key not in original_types:
            node=ET.SubElement(types,'{'+C+'}Override',PartName=key,ContentType=kind);original_types[key]=node
    def add_rel(target):
        i=1
        while 'rId'+str(i) in used:i+=1
        rid='rId'+str(i);used.add(rid)
        ET.SubElement(rels,'{'+R+'}Relationship',Id=rid,Type=D+'/worksheet',Target='/'+target)
        return rid
    for title in (INPUT,DATA,VIEW):
        node,newpath=new_sheets[title]
        if title in old_sheets:
            _,path=old_sheets[title]
            sheet=next(s for s in sheets if s.get('name')==title)
            sheet.set('state',node.get('state','visible'))
        else:
            path='xl/worksheets/ps_'+title.lower().replace(' ','_')+'.xml'
            if path in result:raise ValueError('Finance package path collision')
            sheet=deepcopy(node);sheet.set('sheetId',str(max(int(s.get('sheetId')) for s in sheets)+1))
            sheet.set('{'+D+'}id',add_rel(path));sheets.append(sheet)
        sheet_xml=_xml(generated[newpath])
        for cell in sheet_xml.iter():
            key='s' if cell.tag in ('{'+S+'}c','{'+S+'}row') else 'style' if cell.tag=='{'+S+'}col' else None
            if key and key in cell.attrib:cell.set(key,str(style_map[int(cell.get(key))]))
        result[path]=_bytes(sheet_xml)
        add_type(path,generated_types['/'+newpath].get('ContentType'))
        nr=_relpath(newpath);pr=_relpath(path)
        if pr in original:
            for rel in _xml(original[pr]):
                if rel.get('Type')==D+'/table':
                    owned=_resolve(path,rel.get('Target'))
                    result.pop(owned,None)
                    for typ in list(types):
                        if typ.get('PartName')=='/'+owned:types.remove(typ);original_types.pop('/'+owned,None)
            result.pop(pr,None)
        if nr in generated:
            sheet_rels=_xml(generated[nr])
            for rel in sheet_rels:
                if rel.get('Type')!=D+'/table':
                    # Input notes are plain cells. Arbitrary comments/hyperlinks
                    # are not silently lost by this narrowly supported publisher.
                    raise ValueError('Unsupported relationship on finance sheet: '+rel.get('Type',''))
                tablepath=_resolve(newpath,rel.get('Target'))
                table=_xml(generated[tablepath]);name=table.get('name')
                dest='xl/tables/ps_'+name.lower()+'.xml'
                if dest in original:
                    existing=_xml(original[dest])
                    if existing.get('name')!=name:raise ValueError('Finance table path collision')
                    table.set('id',existing.get('id'))
                else:
                    table.set('id',str(next_table_id));next_table_id+=1
                rel.set('Target','/'+dest);result[dest]=_bytes(table)
                add_type(dest,generated_types['/'+tablepath].get('ContentType'))
            result[pr]=_bytes(sheet_rels)
    old_names=book.find('{'+S+'}definedNames')
    if old_names is None:old_names=ET.SubElement(book,'{'+S+'}definedNames')
    new_names=_xml(generated['xl/workbook.xml']).find('{'+S+'}definedNames')
    if new_names is not None:
        for name in new_names:
            if not name.get('name','').startswith('FF_'):continue
            existing=next((n for n in old_names if n.get('name')==name.get('name')),None)
            if existing is not None:old_names.remove(existing)
            old_names.append(deepcopy(name))
    # Preserve original calcPr exactly. F9/Save policy is owned by the source.
    # Place names before calcPr according to workbook element order.
    book.remove(old_names)
    calc=book.find('{'+S+'}calcPr')
    book.insert(list(book).index(calc) if calc is not None else len(book),old_names)
    result['xl/workbook.xml']=_bytes(book);result['xl/_rels/workbook.xml.rels']=_bytes(rels)
    result['[Content_Types].xml']=_bytes(types)
    with ZipFile(rendered,'w',ZIP_DEFLATED) as z:
        for path,value in result.items():z.writestr(path,value)
