"""Editable Word presentation of the reviewed BOQ report lines."""
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from progress_studio.infrastructure.word.payment_plan import styles, table, add_row
from progress_studio.services.boq_word_report import display_number


def render(path, plan, *, cumulative, check_cancel):
    review = plan.catalog.review
    kind = 'Cumulative' if cumulative else 'Period'
    title = kind + ' BOQ Payment Plan Report'
    doc = Document()
    styles(doc)
    doc.core_properties.title = title
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(1.65)
    section.header.paragraphs[0].text = review.sources[0][0].stem + '    /    BOQ payment planning'
    footer = section.footer.paragraphs[0]
    footer.text = kind + ' BOQ payment plan    /    '
    page = OxmlElement('w:fldSimple')
    page.set(qn('w:instr'), 'PAGE')
    footer._p.append(page)
    doc.add_paragraph(title, 'Title')
    doc.add_paragraph(f'{len(plan.details)} selected rows cover {plan.count} of '
                      f'{len(plan.catalog.entries)} BOQ items across {len(review.projection.periods)} periods.')
    doc.add_paragraph(f'Selected basis {display_number(plan.total.amount)}    /    Project BOQ basis '
                      f'{display_number(plan.catalog.root.amount)}    /    Coverage {display_number(plan.coverage, True)}')
    doc.add_paragraph('Planned amounts use the BOQ source monetary unit. Included detail is counted within its '
                      'selected ancestor. Known contributions can cover part of an incomplete scope.')
    if cumulative:
        doc.add_paragraph('A dash marked Still complete carries the preceding exact completion forward. '
                          'The scope register records the corresponding BOQ basis.')
    doc.add_heading('Reviewed sources', 1)
    for source, revision in review.sources:
        doc.add_paragraph(f'{source.name}   /   SHA-256 {revision}', 'Source note')
    for issue in review.projection.diagnostics:
        if issue.code == 'SourceRevision':
            doc.add_paragraph(issue.detail, 'Source note')
    doc.add_heading('Selected BOQ work', 1)
    register = table(doc, ('BOQ work', 'Scope', 'BOQ basis'), (9, 4.5, 4.2))
    for detail in plan.details:
        check_cancel()
        row = detail.row
        role = 'Included in ' + (detail.ancestor.wbs or detail.ancestor.name) if detail.ancestor else 'Counted'
        add_row(register, ((row.wbs + '  ' + row.name).strip(), role, display_number(row.amount)))
    for index in range(len(review.projection.periods)):
        check_cancel()
        heading = plan.period_title(index)
        doc.add_page_break()
        doc.add_heading(heading, 1)
        grid = table(doc, ('BOQ work  /  ' + review.projection.periods[index], 'Status', 'Requirement', 'Planned amount'),
                     (7.15, 3.75, 2.85, 3.95))
        for position, line in enumerate(plan.lines(index, cumulative)):
            check_cancel()
            total = position >= len(plan.details)
            add_row(grid, line, 'Table total' if total else 'Table text', 'E5F0ED' if total else None,
                    keep=position in (len(plan.details)-1, len(plan.details)))
    doc.save(path)
