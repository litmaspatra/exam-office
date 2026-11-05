# -*- coding: utf-8 -*-
"""
docx_gen.py – Generate a fresh exam .docx from a Paper object.

The document uses Remington Gail / Kruti Dev 010 (falls back to DevLys 010 / Nirmala UI)
with portrait A4, 1-inch margins, empty header/footer, and no table borders
so students start from an authentic exam slate.
"""

from __future__ import annotations

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from paper_loader import Paper
import kruti_converter

KRUTI_FONT = 'Kruti Dev 010'
BASE_SIZE = Pt(14)
HEADING_SIZE = Pt(16)


def _to_remington(text: str) -> str:
    """Convert Unicode Hindi to Remington Gail (Kruti Dev 010) ASCII character codes."""
    if not text:
        return ''
    return kruti_converter.Unicode_to_KrutiDev(text)


def _set_font_properties(style_or_run, name: str, size) -> None:
    """Set ascii, hAnsi, and complex-script fonts so Remington / Kruti Dev renders correctly in Word."""
    rpr = style_or_run.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)
    sz = rpr.find(qn('w:sz'))
    if sz is None:
        sz = OxmlElement('w:sz')
        rpr.append(sz)
    sz.set(qn('w:val'), str(int(size.pt * 2)))
    szcs = rpr.find(qn('w:szCs'))
    if szcs is None:
        szcs = OxmlElement('w:szCs')
        rpr.append(szcs)
    szcs.set(qn('w:val'), str(int(size.pt * 2)))


def generate(paper: Paper, out_path: str) -> str:
    """Write a fresh exam document for *paper* to *out_path* in Remington Gail (Kruti Dev 010)."""
    doc = Document()

    # --- base styles --------------------------------------------------------
    normal = doc.styles['Normal']
    normal.font.name = KRUTI_FONT
    normal.font.size = BASE_SIZE
    _set_font_properties(normal, KRUTI_FONT, BASE_SIZE)

    body_style = doc.styles.add_style('EOBody', WD_STYLE_TYPE.PARAGRAPH)
    body_style.base_style = doc.styles['Normal']
    body_style.font.name = KRUTI_FONT
    body_style.font.size = BASE_SIZE
    _set_font_properties(body_style, KRUTI_FONT, BASE_SIZE)

    head_style = doc.styles.add_style('EOHeading', WD_STYLE_TYPE.PARAGRAPH)
    head_style.base_style = doc.styles['Normal']
    head_style.font.name = KRUTI_FONT
    head_style.font.size = HEADING_SIZE
    head_style.font.bold = True
    _set_font_properties(head_style, KRUTI_FONT, HEADING_SIZE)

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
    head_p = doc.add_paragraph(style='EOHeading')
    head_r = head_p.add_run(_to_remington(d['heading']))
    head_r.font.name = KRUTI_FONT
    head_r.font.size = HEADING_SIZE
    head_r.bold = True
    _set_font_properties(head_r, KRUTI_FONT, HEADING_SIZE)

    for p_text in d.get('paragraphs', []):
        p = doc.add_paragraph(style='EOBody')
        r = p.add_run(_to_remington(p_text))
        r.font.name = KRUTI_FONT
        r.font.size = BASE_SIZE
        _set_font_properties(r, KRUTI_FONT, BASE_SIZE)

    if 'table' in d and d['table']:
        tdata = [d['table']['head']] + d['table']['rows']
        table = doc.add_table(rows=len(tdata), cols=len(tdata[0]))
        for r_idx, row in enumerate(tdata):
            for c_idx, txt in enumerate(row):
                cell = table.rows[r_idx].cells[c_idx]
                cell.paragraphs[0].style = doc.styles['EOBody']
                cell_r = cell.paragraphs[0].add_run(_to_remington(txt))
                cell_r.font.name = KRUTI_FONT
                cell_r.font.size = BASE_SIZE
                _set_font_properties(cell_r, KRUTI_FONT, BASE_SIZE)

    doc.save(out_path)
    return out_path
