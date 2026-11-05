# -*- coding: utf-8 -*-
"""
test_marking.py – Automated tests for ExamOffice paper loading, docx generation,
evaluation/marking engine, error handling, and persistence.
"""

import io
import json
import os
import shutil
import tempfile
import pytest
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_COLOR_INDEX
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from paper_loader import (
    load_all_papers,
    _load_from_folder,
    PaperLoadError,
    Paper,
    Task,
)
import docx_gen
import marker
import store
from exam_office import _safe_grade


@pytest.fixture
def temp_dir():
    d = tempfile.mkdtemp()
    yield d
    shutil.rmtree(d, ignore_errors=True)


def test_paper_manifests_load():
    papers, errors = load_all_papers()
    assert not errors, f"Errors loading papers: {errors}"
    ids = {p.id for p in papers}
    assert 'w-1' in ids
    assert 'w-2' in ids
    for p in papers:
        assert len(p.tasks) == 15
        assert p.max_marks == 25
        assert sum(t.marks for t in p.tasks) == 25


def test_schema_validations(temp_dir):
    valid_manifest = {
        "schema_version": 1,
        "id": "test-paper",
        "title": "Test Paper",
        "duration_seconds": 600,
        "max_marks": 5,
        "pass_marks": 2.0,
        "pass_marks_scst": 1.5,
        "document": {
            "heading": "Heading",
            "paragraphs": ["Para 1"]
        },
        "tasks": [
            {
                "id": "t-1",
                "text": "Task 1",
                "marks": 5,
                "check": {"kind": "align", "para": 0, "val": "justify"}
            }
        ]
    }

    # Helper to test invalid manifest
    def assert_load_fails(data_dict, match_str):
        p_folder = tempfile.mkdtemp(dir=temp_dir)
        with open(os.path.join(p_folder, "paper.json"), "w", encoding="utf-8") as f:
            json.dump(data_dict, f)
        with pytest.raises(PaperLoadError) as exc_info:
            _load_from_folder(p_folder)
        assert match_str in str(exc_info.value)

    # Missing/invalid id
    bad = dict(valid_manifest)
    del bad["id"]
    assert_load_fails(bad, "Invalid or missing 'id'")

    # Wrong schema_version
    bad = dict(valid_manifest, schema_version=99)
    assert_load_fails(bad, "schema_version must be 1")

    # Duplicate task id
    bad = dict(valid_manifest, tasks=[
        {"id": "t-1", "text": "A", "marks": 2, "check": {"kind": "align", "para": 0, "val": "justify"}},
        {"id": "t-1", "text": "B", "marks": 3, "check": {"kind": "align", "para": 0, "val": "justify"}},
    ])
    assert_load_fails(bad, "duplicate task id 't-1'")

    # Total marks mismatch
    bad = dict(valid_manifest, tasks=[
        {"id": "t-1", "text": "A", "marks": 2, "check": {"kind": "align", "para": 0, "val": "justify"}}
    ])
    assert_load_fails(bad, "total marks (2) != max_marks (5)")


def solve_paper_1(path):
    doc = Document(path)
    body = [p for p in doc.paragraphs if p.style.name == 'EOBody']
    heading = next(p for p in doc.paragraphs if p.style.name == 'EOHeading')

    # t-w1-01: justify para 0
    body[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # t-w1-02: first_line_indent 0.5" para 1
    body[1].paragraph_format.first_line_indent = Inches(0.5)

    # t-w1-03: linespacing 1.5 para 2
    body[2].paragraph_format.line_spacing = 1.5

    # t-w1-04: bold para 3
    for r in body[3].runs:
        r.bold = True

    # t-w1-05: heading center + size 16
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in heading.runs:
        r.font.size = Pt(16)

    # t-w1-06: color red para 1
    for r in body[1].runs:
        r.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    # t-w1-07: phrase_italic "जल है तो कल है" (or Remington 'ty gS rks dy gS') in para 3
    full_text = body[3].text
    # Handle either Remington Kruti Dev or Unicode
    phrase = 'ty gS rks dy gS' if 'ty gS rks dy gS' in full_text else 'जल है तो कल है'
    idx = full_text.find(phrase)
    if idx != -1:
        before = full_text[:idx]
        after = full_text[idx + len(phrase):]
        for r in body[3].runs:
            r.text = ''
        if body[3].runs:
            body[3].runs[0].text = before
        else:
            body[3].add_run(before)
        r_italic = body[3].add_run(phrase)
        r_italic.italic = True
        r_italic.bold = True
        r_after = body[3].add_run(after)
        r_after.bold = True

    # t-w1-08: spaceafter 12pt para 2
    body[2].paragraph_format.space_after = Pt(12)

    # t-w1-09: header text "जल संरक्षण"
    doc.sections[0].header.paragraphs[0].text = 'जल संरक्षण'

    # t-w1-10: footer page number
    footer_para = doc.sections[0].footer.paragraphs[0]
    footer_para.text = ''
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), ' PAGE ')
    r_el = OxmlElement('w:r')
    t_el = OxmlElement('w:t')
    t_el.text = '1'
    r_el.append(t_el)
    fld.append(r_el)
    footer_para._p.append(fld)

    # t-w1-11: find_replace "पानी" -> "जल" (or Remington 'ikuh' -> 'ty')
    for p in doc.paragraphs:
        for r in p.runs:
            if 'पानी' in r.text:
                r.text = r.text.replace('पानी', 'जल')
            if 'ikuh' in r.text:
                r.text = r.text.replace('ikuh', 'ty')
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if 'पानी' in r.text:
                            r.text = r.text.replace('पानी', 'जल')
                        if 'ikuh' in r.text:
                            r.text = r.text.replace('ikuh', 'ty')

    # t-w1-12: orientation landscape + margins narrow 0.5"
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11.69), Inches(8.27)
    for attr in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'):
        setattr(sec, attr, Inches(0.5))

    # t-w1-13: para_shading para 2 (light green)
    ppr = body[2]._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '92D050')
    ppr.append(shd)

    # t-w1-14: table head bold + shade yellow
    table = doc.tables[0]
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
        tcpr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'FFFF00')
        tcpr.append(shd)

    # t-w1-15: table borders + cell [3][2] = "500"
    tbl = table._tbl
    tblpr = tbl.tblPr
    if tblpr is None:
        tblpr = OxmlElement('w:tblPr')
        tbl.insert(0, tblpr)
    tblBorders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        tblBorders.append(el)
    tblpr.append(tblBorders)

    cell = table.rows[3].cells[2]
    cell.paragraphs[0].text = ''
    cell.paragraphs[0].add_run('500')

    doc.save(path)


def solve_paper_2(path):
    doc = Document(path)
    body = [p for p in doc.paragraphs if p.style.name == 'EOBody']
    heading = next(p for p in doc.paragraphs if p.style.name == 'EOHeading')

    # t-w2-01: heading bold + color blue 0000FF
    for r in heading.runs:
        r.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)

    # t-w2-02: center para 0
    body[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # t-w2-03: italic para 1
    for r in body[1].runs:
        r.italic = True

    # t-w2-04: linespacing 2.0 para 2
    body[2].paragraph_format.line_spacing = 2.0

    # t-w2-05: indent 0.5" para 3
    body[3].paragraph_format.left_indent = Inches(0.5)

    # t-w2-06: word_double_underline "ऑनलाइन" in para 1
    from paper_loader import load_all_papers
    import kruti_converter
    full = body[1].text
    word_uni = 'ऑनलाइन'
    word_kruti = kruti_converter.Unicode_to_KrutiDev(word_uni)
    word = word_kruti if word_kruti in full else word_uni
    idx = full.find(word)
    if idx != -1:
        before = full[:idx]
        after = full[idx + len(word):]
        for r in body[1].runs:
            r.text = ''
        if body[1].runs:
            body[1].runs[0].text = before
        else:
            body[1].add_run(before)
        r_du = body[1].add_run(word)
        r_du.italic = True
        r_du.font.underline = WD_UNDERLINE.DOUBLE
        r_after = body[1].add_run(after)
        r_after.italic = True

    # t-w2-07: highlight yellow para 3
    for r in body[3].runs:
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    # t-w2-08: page_break_before para 0
    body[0].paragraph_format.page_break_before = True

    # t-w2-09: header text "डिजिटल साक्षरता"
    doc.sections[0].header.paragraphs[0].text = 'डिजिटल साक्षरता'

    # t-w2-10: footer_text_bold "परीक्षा 2026" with "2026" bold
    footer_para = doc.sections[0].footer.paragraphs[0]
    footer_para.text = ''
    footer_para.add_run('परीक्षा ')
    r_bold = footer_para.add_run('2026')
    r_bold.bold = True

    # t-w2-11: find_replace "कम्प्यूटर" -> "कंप्यूटर"
    for p in doc.paragraphs:
        for r in p.runs:
            if 'कम्प्यूटर' in r.text:
                r.text = r.text.replace('कम्प्यूटर', 'कंप्यूटर')
            if 'dEI;wVj' in r.text:
                r.text = r.text.replace('dEI;wVj', 'daI;wVj')
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if 'कम्प्यूटर' in r.text:
                            r.text = r.text.replace('कम्प्यूटर', 'कंप्यूटर')
                        if 'dEI;wVj' in r.text:
                            r.text = r.text.replace('dEI;wVj', 'daI;wVj')

    # t-w2-12: page_color (non-white)
    bg = OxmlElement('w:background')
    bg.set(qn('w:color'), 'ADD8E6')
    doc.element.insert(0, bg)

    # t-w2-13: page_border
    sectPr = doc.sections[0]._sectPr
    pgBorders = OxmlElement('w:pgBorders')
    pgBorders.set(qn('w:offsetFrom'), 'page')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '24')
        el.set(qn('w:color'), '000000')
        pgBorders.append(el)
    sectPr.append(pgBorders)

    # t-w2-14: table head bold + center
    table = doc.tables[0]
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True

    # t-w2-15: cell_text row 1 col 2 = "75%"
    cell = table.rows[1].cells[2]
    cell.paragraphs[0].text = ''
    cell.paragraphs[0].add_run('75%')

    doc.save(path)


def test_paper_1_full_score(temp_dir):
    papers, _ = load_all_papers()
    p1 = next(p for p in papers if p.id == 'w-1')
    docx_path = os.path.join(temp_dir, 'w-1.docx')
    docx_gen.generate(p1, docx_path)

    solve_paper_1(docx_path)
    res = marker.grade(docx_path, p1)

    failed = [f"{tr.task.id}: {tr.hint}" for tr in res.task_results if not tr.passed]
    assert not failed, f"Failed tasks: {failed}"
    assert res.earned_marks == 25


def test_paper_2_full_score(temp_dir):
    papers, _ = load_all_papers()
    p2 = next(p for p in papers if p.id == 'w-2')
    docx_path = os.path.join(temp_dir, 'w-2.docx')
    docx_gen.generate(p2, docx_path)

    solve_paper_2(docx_path)
    res = marker.grade(docx_path, p2)

    failed = [f"{tr.task.id}: {tr.hint}" for tr in res.task_results if not tr.passed]
    assert not failed, f"Failed tasks: {failed}"
    assert res.earned_marks == 25


def test_partial_score(temp_dir):
    papers, _ = load_all_papers()
    p1 = next(p for p in papers if p.id == 'w-1')
    docx_path = os.path.join(temp_dir, 'w-1_partial.docx')
    docx_gen.generate(p1, docx_path)

    # Apply only task 1 (justify) and task 4 (bold)
    doc = Document(docx_path)
    body = [p for p in doc.paragraphs if p.style.name == 'EOBody']
    body[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for r in body[3].runs:
        r.bold = True
    doc.save(docx_path)

    res = marker.grade(docx_path, p1)
    assert 0 < res.earned_marks < 25
    assert res.task_results[0].passed is True   # t-w1-01
    assert res.task_results[1].passed is False  # t-w1-02
    assert res.task_results[3].passed is True   # t-w1-04


def test_wrong_answer(temp_dir):
    papers, _ = load_all_papers()
    p1 = next(p for p in papers if p.id == 'w-1')
    docx_path = os.path.join(temp_dir, 'w-1_wrong.docx')
    docx_gen.generate(p1, docx_path)

    # Set wrong color (green instead of red) on paragraph 1
    doc = Document(docx_path)
    body = [p for p in doc.paragraphs if p.style.name == 'EOBody']
    for r in body[1].runs:
        r.font.color.rgb = RGBColor(0x00, 0xFF, 0x00)
    doc.save(docx_path)

    res = marker.grade(docx_path, p1)
    task_6 = next(tr for tr in res.task_results if tr.task.id == 't-w1-06')
    assert task_6.passed is False
    assert "लाल (FF0000) नहीं मिला" in task_6.hint


def test_missing_docx(temp_dir):
    papers, _ = load_all_papers()
    p1 = papers[0]
    non_existent = os.path.join(temp_dir, 'does_not_exist.docx')
    res = _safe_grade(non_existent, p1, max_retries=1, delay=0.01)
    assert res is None


def test_corrupt_docx(temp_dir):
    papers, _ = load_all_papers()
    p1 = papers[0]
    corrupt_path = os.path.join(temp_dir, 'corrupt.docx')
    with open(corrupt_path, 'wb') as f:
        f.write(b"this is corrupt data not a zip file")
    res = _safe_grade(corrupt_path, p1, max_retries=1, delay=0.01)
    assert res is None


def test_result_persistence(temp_dir):
    papers, _ = load_all_papers()
    p1 = papers[0]
    docx_path = os.path.join(temp_dir, 'sample.docx')
    docx_gen.generate(p1, docx_path)

    res = marker.grade(docx_path, p1)
    attempt_dir = temp_dir
    record = {
        'attempt_id': 'test-att-001',
        'paper_id': p1.id,
        'paper_title': p1.title,
        'mode': 'practice',
        'earned_marks': res.earned_marks,
        'total_marks': res.total_marks,
        'tasks': [{'id': tr.task.id, 'passed': tr.passed, 'earned': tr.earned} for tr in res.task_results],
        'time_used': '05:30',
        'timestamp': '2026-11-03 14:00',
    }
    store.save_result(attempt_dir, record)
    assert os.path.isfile(store.result_json_path(attempt_dir))

    loaded = store.load_result(attempt_dir)
    assert loaded is not None
    assert loaded['paper_id'] == p1.id
    assert loaded['mode'] == 'practice'
    assert loaded['earned_marks'] == res.earned_marks
    assert len(loaded['tasks']) == len(p1.tasks)
