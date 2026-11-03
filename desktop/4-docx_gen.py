# -*- coding: utf-8 -*-
"""Generate a fresh exam .docx for a paper: heading + 4 paragraphs + 3x4 table,
header/footer empty, portrait, normal margins, no table borders, base font 12pt
(Nirmala UI / Mangal for Devanagari complex script)."""
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE_FONT = 'Nirmala UI'   # Windows Devanagari UI font; falls back to Mangal on old systems
BASE_SIZE = Pt(12)
HEADING_SIZE = Pt(14)


def _set_cs_font(style_or_run, name, size):
    """Set both latin and complex-script font so Hindi renders in Word."""
    rpr = style_or_run.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rfonts.set(qn(attr), name)
    szcs = rpr.find(qn('w:szCs'))
    if szcs is None:
        szcs = OxmlElement('w:szCs')
        rpr.append(szcs)
    szcs.set(qn('w:val'), str(int(size.pt * 2)))


def generate(paper, out_path):
    doc = Document()

    # base style
    normal = doc.styles['Normal']
    normal.font.name = BASE_FONT
    normal.font.size = BASE_SIZE
    _set_cs_font(normal, BASE_FONT, BASE_SIZE)

    body_style = doc.styles.add_style('EOBody', WD_STYLE_TYPE.PARAGRAPH)
    body_style.base_style = doc.styles['Normal']
    body_style.font.name = BASE_FONT
    body_style.font.size = BASE_SIZE
    _set_cs_font(body_style, BASE_FONT, BASE_SIZE)

    head_style = doc.styles.add_style('EOHeading', WD_STYLE_TYPE.PARAGRAPH)
    head_style.base_style = doc.styles['Normal']
    head_style.font.name = BASE_FONT
    head_style.font.size = HEADING_SIZE
    head_style.font.bold = True
    _set_cs_font(head_style, BASE_FONT, HEADING_SIZE)

    # section: portrait A4, normal (1 inch) margins, empty header/footer
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    for attr in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'):
        setattr(sec, attr, Inches(1))
    sec.different_first_page_header_footer = False
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    sec.header.paragraphs[0].text = ''
    sec.footer.paragraphs[0].text = ''

    # heading + paragraphs
    doc.add_paragraph(paper['heading'], style='EOHeading')
    for p in paper['paras']:
        doc.add_paragraph(p, style='EOBody')

    # table: head + 3 rows, no borders
    tdata = [paper['table']['head']] + paper['table']['rows']
    table = doc.add_table(rows=len(tdata), cols=len(tdata[0]))
    for r, row in enumerate(tdata):
        for c, txt in enumerate(row):
            cell = table.rows[r].cells[c]
            cell.paragraphs[0].style = doc.styles['EOBody']
            cell.paragraphs[0].add_run(txt)

    doc.save(out_path)
    return out_path
