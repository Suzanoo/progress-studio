"""Dependency-free saved-value snapshots; no Payment or Requirement calculation."""
from copy import copy, deepcopy
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell, Cell
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import get_column_letter
from openpyxl.styles import Protection, Alignment
from progress_studio.config.workbook_protection import WORKBOOK_SHEET_PASSWORD
from progress_studio.infrastructure.excel.payment_workbook import PaymentWorkbookSnapshotter

SHEETS = ('Payment Input', 'Payment')
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


@dataclass(frozen=True)
class SnapshotInfo:
    sheet: str
    cells: int
    formulas: int
    errors: int


def inspect_snapshots(path, cancelled=lambda: False):
    """Validate actual stored cells, never trust inflated worksheet dimensions.

    Empty string formula caches (t=str, empty v) are valid blanks. A missing
    numeric cache is unavailable, not zero. Cache freshness cannot be certified.
    """
    result = []
    with ZipFile(path) as archive:
        sheet_map = PaymentWorkbookSnapshotter._sheet_map(archive)
        for name in SHEETS:
            if cancelled():raise InterruptedError('BOQ export cancelled')
            if name not in sheet_map:raise ValueError(f"Snapshot source worksheet missing: {name}")
            root = ET.fromstring(archive.read(sheet_map[name]))
            formulas = errors = cells = 0
            for cell in root.findall('./s:sheetData/s:row/s:c', NS):
                cells += 1
                if cells % 2048 == 0 and cancelled():raise InterruptedError('BOQ export cancelled')
                if cell.find('s:f', NS) is not None:
                    formulas += 1
                    value = cell.find('s:v', NS)
                    if value is None or (value.text is None and cell.get('t') != 'str'):
                        raise ValueError(f"Missing saved formula value: {name}!{cell.get('r')}. Recalculate and save the source workbook, then review again.")
                if cell.get('t') == 'e':errors += 1
            result.append(SnapshotInfo(name, cells, formulas, errors))
    return tuple(result)


def _selected_package(path):
    """Reuse source styles/media without loading unrelated worksheets into RAM."""
    out = BytesIO()
    with ZipFile(path) as src, ZipFile(out, 'w', ZIP_DEFLATED) as dst:
        root = ET.fromstring(src.read('xl/workbook.xml'))
        sheets = root.find('s:sheets', NS)
        for sheet in list(sheets):
            if sheet.get('name') not in SHEETS:sheets.remove(sheet)
        for tag in ('definedNames', 'externalReferences'):
            node = root.find('s:' + tag, NS)
            if node is not None:root.remove(node)
        for view in root.findall('s:bookViews/s:workbookView', NS):view.set('activeTab','0')
        for info in src.infolist():
            if info.filename == 'xl/calcChain.xml':continue
            content = ET.tostring(root) if info.filename == 'xl/workbook.xml' else src.read(info.filename)
            dst.writestr(info, content)
    out.seek(0)
    return out


def _local_presentation_formula(formula, sheet):
    # Presentation CF may refer to frozen helper values on this same sheet only.
    formula = formula.replace("'" + sheet.replace("'", "''") + "'!", '').replace(sheet + '!', '')
    for token in Tokenizer('=' + formula.lstrip('=')).items:
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            if not re.fullmatch(r'\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?', token.value, re.I):
                raise ValueError(f'Unsupported snapshot formatting dependency: {sheet}: {token.value}')
        if token.type == 'FUNC' and token.subtype == 'OPEN':
            if token.value.upper() not in {'IF(', 'IFERROR(', 'AND(', 'OR(', 'NOT(', 'ISNUMBER(', 'SUM(', 'COUNT(', 'COUNTIF(', 'COUNTIFS(', 'COLUMN(', 'ROW(', 'ISBLANK('}:
                raise ValueError(f'Unsupported snapshot formatting function: {token.value}')
    return formula


def _freeze_chart(chart):
    """Retain chart geometry/format with literal cached series, never source links."""
    chart = deepcopy(chart)
    for subchart in chart._charts:
        for series in subchart.series:
            for attr in ('val', 'cat', 'xVal', 'yVal', 'bubbleSize'):
                data = getattr(series, attr, None)
                if data is None:continue
                for ref_name, cache_name, literal_name in (('numRef','numCache','numLit'),('strRef','strCache','strLit')):
                    ref = getattr(data, ref_name, None)
                    if ref is not None:
                        cache = getattr(ref, cache_name)
                        if cache is None:raise ValueError('Missing saved chart values; recalculate and save source')
                        setattr(data, literal_name, deepcopy(cache));setattr(data, ref_name, None)
            title = series.tx
            if title is not None and title.strRef is not None:
                cache = title.strRef.strCache
                if cache is None or not cache.pt:raise ValueError('Missing saved chart title')
                title.v = cache.pt[0].v;title.strRef = None
    return chart


def append_payment_snapshots(target, source_path, source_hash, cancelled=lambda: False):
    info = inspect_snapshots(source_path, cancelled)
    # Reject unsupported DrawingML shapes rather than silently dropping content.
    with ZipFile(source_path) as package:
        sheet_map = PaymentWorkbookSnapshotter._sheet_map(package)
        from posixpath import normpath, dirname, join
        for name in SHEETS:
            rels = PaymentWorkbookSnapshotter._worksheet_rels_part(sheet_map[name])
            if rels not in package.namelist():continue
            for rel in ET.fromstring(package.read(rels)):
                if rel.get('Type','').endswith('/drawing'):
                    part = normpath(join(dirname(sheet_map[name]),rel.get('Target'))).lstrip('/')
                    drawing = ET.fromstring(package.read(part))
                    if any(e.tag.rsplit('}',1)[-1] in {'sp','grpSp','cxnSp','contentPart'} for e in drawing.iter()):
                        raise ValueError(f'Unsupported static drawing in {name}; snapshot not published')
    payload = _selected_package(source_path)
    source = load_workbook(payload, data_only=True, keep_links=False)
    try:
        for name in SHEETS:
            if cancelled():raise InterruptedError('BOQ export cancelled')
            original = source[name];sheet = target.create_sheet(name)
            # openpyxl worksheet copy cannot cross workbooks. Register each style
            # once in the target, then reuse target-owned style indices.
            styles = {}
            def style(obj):
                key = tuple(obj._style) if obj._style else ()
                if key not in styles:
                    cell = Cell(sheet, row=1, column=1)
                    for attr in ('font','fill','border','alignment','number_format'):
                        setattr(cell,attr,copy(getattr(obj,attr)))
                    cell.protection = Protection(locked=True)
                    styles[key] = copy(cell._style)
                return styles[key]
            # Preserve sparse geometry; do not iterate max_row x max_column.
            for index, src in enumerate(original._cells.values()):
                if index % 2048 == 0 and cancelled():raise InterruptedError('BOQ export cancelled')
                if isinstance(src,MergedCell):continue
                dst = sheet.cell(src.row,src.column)
                dst._style = style(src)
                dst.value = src.value
                dst.data_type = src.data_type
                if src.comment:dst.comment = copy(src.comment)
                # Links and input validations intentionally do not survive.
            for kind in ('row_dimensions','column_dimensions'):
                for key, dim in getattr(original,kind).items():
                    copied=copy(dim);copied.parent=sheet
                    if dim.has_style:copied._style=style(dim)
                    copied.protection=Protection(locked=True)
                    # Locked snapshots cannot reliably toggle native outline in Excel.
                    # Expand grouped rows so all saved Activity detail is readable.
                    if kind == 'row_dimensions' and copied.outlineLevel:
                        copied.hidden = False
                    if kind == 'row_dimensions':copied.collapsed = False
                    getattr(sheet,kind)[key]=copied
            for merged in original.merged_cells.ranges:sheet.merge_cells(str(merged))
            for attr in ('sheet_format','sheet_properties','page_margins','page_setup','print_options','row_breaks','col_breaks','auto_filter'):
                setattr(sheet,attr,deepcopy(getattr(original,attr)))
            sheet.sheet_properties.codeName = None
            # topLeftCell can be a scrolled origin, not the split coordinates.
            # Copy the actual pane rather than feeding that origin to freeze_panes.
            sheet.sheet_view.pane = deepcopy(original.sheet_view.pane)
            sheet.sheet_view.selection = deepcopy(original.sheet_view.selection)
            pane = sheet.sheet_view.pane
            header = next((c.row for c in original._cells.values()
                           if c.row <= 30 and str(c.value).strip().lower() == 'activity id'), None)
            if pane is not None and pane.state in ('frozen','frozenSplit') and header:
                # Expandable source may freeze almost its entire collapsed tree.
                # In an expanded static report retain readable identifier/header panes.
                if (pane.ySplit or 0) > header + 2:pane.ySplit = header
                pane.topLeftCell = f'{get_column_letter(int(pane.xSplit or 0)+1)}{int(pane.ySplit or 0)+1}'
                for selection in sheet.sheet_view.selection:
                    if selection.pane == pane.activePane:
                        selection.activeCell=pane.topLeftCell;selection.sqref=pane.topLeftCell
            sheet.sheet_view.showGridLines = original.sheet_view.showGridLines
            sheet.sheet_view.zoomScale = original.sheet_view.zoomScale
            sheet.print_area = str(original.print_area).replace("'"+name+"'!",'') if original.print_area else None
            sheet.print_title_rows=original.print_title_rows;sheet.print_title_cols=original.print_title_cols
            sheet.conditional_formatting=deepcopy(original.conditional_formatting)
            for rules in sheet.conditional_formatting._cf_rules.values():
                for rule in rules:
                    rule.formula=[_local_presentation_formula(f,name) for f in rule.formula]
                    for block in (rule.colorScale,rule.dataBar,rule.iconSet):
                        if block is not None:
                            for threshold in block.cfvo:
                                if threshold.type=='formula':threshold.val=_local_presentation_formula(threshold.val,name)
            for image in original._images:
                copied=Image(BytesIO(image._data()));copied.anchor=deepcopy(image.anchor);sheet.add_image(copied)
            for chart in original._charts:sheet.add_chart(_freeze_chart(chart),deepcopy(chart.anchor))
            title=original['A1'].value or name
            sheet['A1']=f'SNAPSHOT — {title}. Read-only saved values; edit source and re-export. Cache freshness unverified.'
            sheet['A1'].alignment=Alignment(wrap_text=True,vertical='center')
            if name=='Payment Input' and not any(c.value is not None for c in list(sheet[1])[1:8]) and not any(r.min_row==1 for r in sheet.merged_cells.ranges):
                sheet.merge_cells('A1:H1')
            sheet.row_dimensions[1].height=max(original.row_dimensions[1].height or 15,42)
            sheet['A1'].comment=Comment(f'Source: {Path(source_path).name}\nSHA-256: {source_hash}\nWorksheet: {name}\nSaved cached values are preserved without recalculating Payment. Recalculate/save source before export. See Amount for revision diagnostics.','Progress Studio')
            sheet.oddHeader.center.text='Read-only Payment snapshot'
            sheet.protection.sheet=True;sheet.protection.password=WORKBOOK_SHEET_PASSWORD
            sheet.protection.selectLockedCells=False
            sheet.sheet_state='visible'
        # Keep reports prominent and internal data last.
        target._sheets.sort(key=lambda s: ('Percentage','Amount','BOQ Mapping','Activity Mapping','Payment Input','Payment','_BP_Data').index(s.title))
    finally:
        source.close();payload.close()
    return info
