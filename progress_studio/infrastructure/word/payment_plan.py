"""Native editable Word presentation of one accepted Payment Plan selection."""
from collections import Counter

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


STATES = {'completed': 'Complete', 'adjusted_complete': 'Adjusted after completion',
          'reopened': 'Reopened', 'recompleted': 'Completed again', 'unresolved': 'Incomplete'}


def fmt(value, percent=False):
    if value is None:
        return 'Unavailable'
    return format(value, '.2%' if percent else ',.2f')


def identify(row):
    return row.activities[0] if row.activity else ('WBS ' + row.code if row.code else 'Project')


def styles(document):
    for name, size, bold, color in (
        ('Normal', 10, False, '253A45'), ('Title', 24, True, '172D38'),
        ('Heading 1', 15, True, '225C57'), ('Table text', 9, False, '253A45'),
        ('Table heading', 9, True, 'FFFFFF'), ('Table total', 10, True, '123D3A'),
        ('Source note', 9, False, '465D68')):
        style = document.styles[name] if name in document.styles else document.styles.add_style(
            name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = 'Calibri'
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = RGBColor.from_string(color)
        fonts = style.element.get_or_add_rPr().rFonts
        fonts.set(qn('w:cs'), 'Tahoma')
        fonts.set(qn('w:eastAsia'), 'Tahoma')
        para = style.paragraph_format
        para.space_after = Pt(4 if name in ('Normal','Source note') else 2)
        if name.startswith('Table'):
            para.space_before = Pt(3)
        if name == 'Heading 1':
            para.space_before = Pt(13)
            para.keep_with_next = True
    title_properties = document.styles['Title'].element.get_or_add_pPr()
    border = title_properties.find(qn('w:pBdr'))
    if border is not None:
        title_properties.remove(border)


def table(doc, labels, widths):
    t = doc.add_table(rows=0, cols=len(labels))
    t.autofit = False
    for column, width in zip(t.columns, widths):
        column.width = Cm(width)
    row = add_row(t, labels, 'Table heading', '28515B', keep=True)
    row._tr.get_or_add_trPr().append(OxmlElement('w:tblHeader'))
    return t


def add_row(t, cells, style='Table text', color=None, keep=False, indent=False):
    row = t.add_row()
    row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
    for j, (cell, value) in enumerate(zip(row.cells, cells)):
        cell.width = t.columns[j].width
        if color:
            fill = OxmlElement('w:shd')
            fill.set(qn('w:fill'), color)
            cell._tc.get_or_add_tcPr().append(fill)
        p = cell.paragraphs[0]
        p.style = style
        p.paragraph_format.keep_with_next = keep
        if j:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif indent:
            p.paragraph_format.left_indent = Cm(.35)
        p.add_run(value)
    return row


def render(path, source, selection, *, cumulative, check_cancel):
    kind = 'Cumulative' if cumulative else 'Period'
    title = kind + ' Payment Plan Report'
    doc = Document()
    styles(doc)
    doc.core_properties.title = title
    doc.core_properties.subject = source.project
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(1.9)
    section.top_margin, section.bottom_margin = Cm(2.0), Cm(1.85)
    section.header_distance = section.footer_distance = Cm(.9)
    section.header.paragraphs[0].text = source.project + '   /   Payment planning'
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run(kind + ' payment plan   •   Page ')
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), 'PAGE')
    footer._p.append(field)
    doc.add_paragraph(title, 'Title')
    scope = 'Full project' if selection.full_project else 'Selected works'
    doc.add_paragraph(f'{source.project}    •    {len(source.periods)} payment periods    •    {scope}')
    doc.add_paragraph('Planned requirements for the selected work. Amounts use the source workbook’s monetary units. '
                      'Certification and actual disbursement are outside this report.')
    doc.add_heading('Selected works', 1)
    doc.add_paragraph(f'{len(selection.rows)} reporting rows cover {len(selection.selected.activities)} '
                      f'of {len(selection.project.activities)} project activities.')
    lookup = {r.key: r for r in source.rows}
    wbs_counts = Counter(r.code for r in source.rows if not r.activity)

    def identifier(row):
        if not row.activity and row.parent and wbs_counts[row.code] > 1:
            return row.key
        return identify(row)

    def label(row):
        text = identifier(row) + '   ' + row.name
        parent = selection.included_in.get(row.key)
        if parent:
            text += '\nIncluded in ' + identifier(lookup[parent]) + '   ' + lookup[parent].name
        return text
    register = table(doc, ['Reporting row', 'Amount basis'], [12.9, 4.3])
    for row in selection.rows:
        check_cancel()
        add_row(register, [label(row), fmt(selection.row_values[row.key].basis)],
                indent=row.key in selection.included_in)
    add_row(register, ['Selected basis', fmt(selection.selected.basis)], 'Table total', 'DFEDE9')
    doc.add_paragraph('Project basis   ' + fmt(selection.project.basis), 'Source note')
    doc.add_paragraph('Explanatory rows marked “Included in” contribute through their selected ancestor. '
                      'Totals count each Activity once and use source precision before display rounding.')
    doc.add_paragraph(('Completed detail is retained in the cumulative total. A later adjustment appears '
                       'in the period it occurs.' if cumulative else
                       'Signed amounts show movement since the preceding period. A decrease is identified even '
                       'when it rounds to zero.') + ' The completion display threshold is 99.99%.')
    doc.add_paragraph(f'Source   {source.filename}\nMonetary source   {source.monetary_origin}\n'
                      f'SHA-256   {source.revision}', 'Source note')
    for issue in source.diagnostics:
        doc.add_paragraph(issue, 'Source note')
    for j, period in enumerate(source.periods):
        check_cancel()
        date = period.eligible.strftime('%d %B %Y') if period.eligible else 'Date unavailable'
        heading = doc.add_heading(f'{period.id}    {date}', 1)
        heading.paragraph_format.page_break_before = j == 0
        note = doc.add_paragraph('Planned eligible date   •   ' + kind + ' requirement', 'Source note')
        note.paragraph_format.keep_with_next = True
        for issue in period.issues:
            doc.add_paragraph(issue, 'Source note')
        t = table(doc, ['Work item', kind + ' %', 'Planned amount'], [10.3,3,3.9])
        last = None
        for row in selection.rows:
            v = selection.row_values[row.key].values[j]
            if not (v.cumulative_visible if cumulative else v.period_visible):
                continue
            check_cancel()
            notes = [STATES[v.transition]] if v.transition in STATES else []
            if v.correction:
                notes.append('Decrease')
            name = label(row)
            if notes:
                name += '\n' + ' · '.join(notes)
            last = add_row(t, [name, fmt(v.display_fraction if cumulative else v.delta, True),
                               fmt(v.money if cumulative else v.movement)], indent=row.key in selection.included_in)
        if last is None:
            last = add_row(t, ['No current detail rows.' if cumulative else 'No reportable movement.', '', ''])
        for cell in last.cells:
            cell.paragraphs[0].paragraph_format.keep_with_next = True
        if cumulative and selection.retained_money[j]:
            add_row(t, ['Previously completed work retained', '', fmt(selection.retained_money[j])],
                    color='F0F5F4', keep=True)
        v = selection.selected.values[j]
        add_row(t, ['Selected ' + kind.lower() + ' total', fmt(v.display_fraction if cumulative else v.delta, True),
                    fmt(v.money if cumulative else v.movement)], 'Table total', 'DFEDE9')
        if v.money is None or v.fraction is None:
            doc.add_paragraph(f'Known activity basis {fmt(v.known_basis)} '
                              f'({v.known_leaves}/{len(selection.selected.activities)} activities). '
                              f'Known cumulative contribution {fmt(v.known_money)}.', 'Source note')
    check_cancel()
    doc.save(path)
