"""Append/replace only BP-4's sheet, reusing the accepted package/style merger."""
from copy import deepcopy
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from progress_studio.infrastructure.excel.finance_package import (
    _xml, _bytes, _sheets, _merge_styles, _relpath, S, R, D, C)
from progress_studio.infrastructure.excel.boq_requirement_input import SHEET


def merge_input(source, rendered):
    with ZipFile(source) as z: original = {n:z.read(n) for n in z.namelist()}
    with ZipFile(rendered) as z: generated = {n:z.read(n) for n in z.namelist()}
    result = dict(original)
    result['xl/styles.xml'], styles = _merge_styles(original['xl/styles.xml'], generated['xl/styles.xml'])
    old, new = _sheets(original), _sheets(generated)
    book = _xml(original['xl/workbook.xml']);sheets = book.find('{'+S+'}sheets')
    rels = _xml(original['xl/_rels/workbook.xml.rels']);types = _xml(original['[Content_Types].xml'])
    node, newpath = new[SHEET]
    if SHEET in old:
        _, path = old[SHEET]
        if _relpath(path) in original:
            raise ValueError('Unsupported objects/relationships on BOQ Requirement Input; preserve source')
    else:
        path = 'xl/worksheets/ps_boq_requirement_input.xml'
        if path in original:raise ValueError('BOQ input package path collision')
        used = {r.get('Id') for r in rels};i=1
        while 'rId'+str(i) in used:i+=1
        rid='rId'+str(i)
        ET.SubElement(rels,'{'+R+'}Relationship',Id=rid,Type=D+'/worksheet',Target='/'+path)
        sheet=deepcopy(node);sheet.set('sheetId',str(max(int(s.get('sheetId')) for s in sheets)+1));sheet.set('{'+D+'}id',rid);sheets.append(sheet)
        ET.SubElement(types,'{'+C+'}Override',PartName='/'+path,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
    root = _xml(generated[newpath])
    # Native Excel may store input labels/identities in a shared string table.
    # Inline those strings so Mapping's regenerated package needs no foreign SST.
    shared = list(_xml(generated['xl/sharedStrings.xml'])) if 'xl/sharedStrings.xml' in generated else []
    for cell in root.iter('{'+S+'}c'):
        if cell.get('t') == 's':
            value = cell.find('{'+S+'}v')
            if value is None or not value.text or not 0 <= int(value.text) < len(shared):
                raise ValueError('Invalid shared string in BOQ Requirement Input')
            inline = deepcopy(shared[int(value.text)]);inline.tag = '{'+S+'}is'
            cell.remove(value);cell.set('t','inlineStr');cell.append(inline)
    for cell in root.iter():
        key='s' if cell.tag in ('{'+S+'}c','{'+S+'}row') else 'style' if cell.tag=='{'+S+'}col' else None
        if key and key in cell.attrib:cell.set(key,str(styles[int(cell.get(key))]))
    result[path] = _bytes(root)
    names=book.find('{'+S+'}definedNames')
    if names is None:
        names=ET.Element('{'+S+'}definedNames')
        calc=book.find('{'+S+'}calcPr')
        book.insert(list(book).index(calc) if calc is not None else len(book),names)
    local_id=str(next(i for i,sheet in enumerate(sheets) if sheet.get('name')==SHEET))
    for name in list(names):
        if name.get('name')=='_xlnm.Print_Titles' and name.get('localSheetId')==local_id:names.remove(name)
    title=ET.SubElement(names,'{'+S+'}definedName',name='_xlnm.Print_Titles',localSheetId=local_id)
    title.text=f"'{SHEET}'!$1:$8,'{SHEET}'!$D:$H"

    # Preserve every existing sheet/cache/drawing and original calculation policy.
    result['xl/workbook.xml']=_bytes(book);result['xl/_rels/workbook.xml.rels']=_bytes(rels)
    result['[Content_Types].xml']=_bytes(types)
    with ZipFile(rendered,'w',ZIP_DEFLATED) as z:
        for path,value in result.items():z.writestr(path,value)


def preserve_input(source, target):
    """Mapping Save As carries the canonical input; later review validates conflicts."""
    from pathlib import Path
    import tempfile
    import shutil
    import os
    with ZipFile(source) as z:
        if SHEET not in _sheets({n:z.read(n) for n in ('xl/workbook.xml','xl/_rels/workbook.xml.rels')}):
            return
    fd,name=tempfile.mkstemp(prefix='.bp4-preserve-',suffix='.xlsx',dir=Path(target).parent)
    os.close(fd);temp=Path(name)
    try:
        shutil.copy2(source,temp)
        merge_input(target,temp)
        os.replace(temp,target)
    finally:
        temp.unlink(missing_ok=True)
