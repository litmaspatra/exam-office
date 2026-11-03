# -*- coding: utf-8 -*-
"""
docx_gen.py – Generate a fresh exam .docx from a Paper object.

The document uses Nirmala UI (Windows Devanagari UI font; falls back to Mangal)
with portrait A4, 1-inch margins, empty header/footer, and no table borders
so students start from a clean slate.
"""

from __future__ import annotations

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from paper_loader import Paper

BASE_FONT = 'Nirmala UI'   # Falls back to Mangal on older Windows
BASE_SIZE = Pt(12)
HEADING_SIZE = Pt(14)


def _set_cs_font(style_or_run, name: str, size) -> None:
    """Set both latin and complex-script font so Devanagari renders correctly in Word."""
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


def generate(paper: Paper, out_path: str) -> str:
    """Write a fresh exam document for *paper* to *out_path* and return it."""
    doc = Document()

    # --- base styles --------------------------------------------------------
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

    # --- section: portrait A4, 1-inch margins, empty header/footer ----------
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    for attr in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'):
        setattr(sec, attr, Inches(1))
    sec.different_first_page_header_footer = False
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    sec.header.paragraphs[0].text = ''
    sec.footer.paragraphs[0].text = ''

    # --- content: heading, paragraphs, table --------------------------------
    d = paper.document
    doc.add_paragraph(d['heading'], style='EOHeading')
    for p in d.get('paragraphs', []):
        doc.add_paragraph(p, style='EOBody')

    if 'table' in d and d['table']:
        tdata = [d['table']['head']] + d['table']['rows']
        table = doc.add_table(rows=len(tdata), cols=len(tdata[0]))
        for r, row in enumerate(tdata):
            for c, txt in enumerate(row):
                cell = table.rows[r].cells[c]
                cell.paragraphs[0].style = doc.styles['EOBody']
                cell.paragraphs[0].add_run(txt)

    doc.save(out_path)
    return out_path
