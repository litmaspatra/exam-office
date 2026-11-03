# -*- coding: utf-8 -*-
"""
marker.py – Auto-marking engine for ExamOffice Word efficiency papers.

Parses a saved .docx with python-docx (+ raw XML where needed) and evaluates
each task's check dict against the document. Works with Paper objects loaded
by paper_loader.py.
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_UNDERLINE
from docx.enum.section import WD_ORIENT
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

from paper_loader import Paper, Task


# ── Helpers ───────────────────────────────────────────────────────────────────

def _norm(s: Any) -> str:
    return ' '.join((s or '').split())


def _runs(p):
    return [r for r in p.runs if r.text]


def _all_runs(p, pred) -> bool:
    rs = _runs(p)
    return bool(rs) and all(pred(r) for r in rs)


# ── Result dataclasses ────────────────────────────────────────────────────────

@dataclass
class TaskResult:
    task: Task
    passed: bool
    earned: int
    hint: str  # shown when failed


@dataclass
class GradeResult:
    paper: Paper
    total_marks: int           # max possible
    earned_marks: int          # actually earned
    task_results: list[TaskResult]
    ungraded: bool = False     # True if the document could not be read
    ungraded_reason: str = ''


# ── Grader class ─────────────────────────────────────────────────────────────

class Grader:
    """
    Evaluates a saved .docx against a paper's task list.

    Accepts either a file path (str) or an io.BytesIO object.
    """

    def __init__(self, docx_source: str | io.BytesIO, paper: Paper) -> None:
        self.doc = Document(docx_source)
        self.paper = paper
        self.sec = self.doc.sections[0]
        self.heading = next(
            (p for p in self.doc.paragraphs if p.style.name == 'EOHeading'), None
        )
        self.body = [p for p in self.doc.paragraphs if p.style.name == 'EOBody']
        self.table = self.doc.tables[0] if self.doc.tables else None
        self.fr = paper.find_replace

    # ── Text-unchanged guard ──────────────────────────────────────────────────

    def _text_ok(self, p, orig: str) -> bool:
        t = _norm(p.text)
        if t == _norm(orig):
            return True
        if self.fr and t == _norm(orig.replace(self.fr['find'], self.fr['replace'])):
            return True
        return False

    def _para(self, i: int):
        return self.body[i] if i < len(self.body) else None

    def _para_task(self, i: int, pred, orig_idx: int | None = None) -> bool:
        p = self._para(i)
        if p is None:
            return False
        oi = orig_idx if orig_idx is not None else i
        paras = self.paper.document.get('paragraphs', [])
        if oi >= len(paras):
            return False
        return pred(p) and self._text_ok(p, paras[oi])

    # ── Individual check dispatch ─────────────────────────────────────────────

    def _check(self, check: dict) -> bool:
        k = check['kind']

        if k == 'align':
            want = (WD_ALIGN_PARAGRAPH.CENTER if check['val'] == 'center'
                    else WD_ALIGN_PARAGRAPH.JUSTIFY)
            return self._para_task(check['para'], lambda p: p.alignment == want)

        if k == 'bold':
            return self._para_task(check['para'], lambda p: _all_runs(p, lambda r: r.bold))

        if k == 'italic':
            return self._para_task(check['para'], lambda p: _all_runs(p, lambda r: r.italic))

        if k == 'underline':
            return self._para_task(check['para'], lambda p: _all_runs(p, lambda r: r.underline))

        if k == 'linespacing':
            def ls(p):
                v = p.paragraph_format.line_spacing
                try:
                    return v is not None and abs(float(v) - check['val']) < 0.01
                except (TypeError, ValueError):
                    return False
            return self._para_task(check['para'], ls)

        if k == 'indent':
            return self._para_task(
                check['para'],
                lambda p: (p.paragraph_format.left_indent is not None
                           and abs(p.paragraph_format.left_indent.inches - check['val']) < 0.06)
            )

        if k == 'first_line_indent':
            return self._para_task(
                check['para'],
                lambda p: (p.paragraph_format.first_line_indent is not None
                           and abs(p.paragraph_format.first_line_indent.inches - check['val']) < 0.06)
            )

        if k == 'spaceafter':
            return self._para_task(
                check['para'],
                lambda p: (p.paragraph_format.space_after is not None
                           and abs(p.paragraph_format.space_after.pt - check['val']) < 1.0)
            )

        if k == 'size':
            return self._para_task(
                check['para'],
                lambda p: _all_runs(
                    p, lambda r: r.font.size is not None and abs(r.font.size.pt - check['val']) < 0.5)
            )

        if k == 'color':
            vals = [v.upper() for v in check['val']]
            return self._para_task(
                check['para'],
                lambda p: _all_runs(
                    p, lambda r: (r.font.color is not None
                                  and r.font.color.rgb is not None
                                  and str(r.font.color.rgb).upper() in vals))
            )

        if k == 'highlight':
            return self._para_task(
                check['para'],
                lambda p: _all_runs(p, lambda r: r.font.highlight_color == WD_COLOR_INDEX.YELLOW)
            )

        if k == 'heading_bold':
            return (self.heading is not None
                    and _all_runs(self.heading, lambda r: r.bold)
                    and self._text_ok(self.heading, self.paper.document.get('heading', '')))

        if k == 'heading_size':
            return (self.heading is not None
                    and _all_runs(self.heading, lambda r: r.font.size is not None
                                  and abs(r.font.size.pt - check['val']) < 0.5)
                    and self._text_ok(self.heading, self.paper.document.get('heading', '')))

        if k == 'heading_color':
            vals = [v.upper() for v in check['val']]
            return (self.heading is not None
                    and _all_runs(self.heading, lambda r: (r.font.color is not None
                                                           and r.font.color.rgb is not None
                                                           and str(r.font.color.rgb).upper() in vals))
                    and self._text_ok(self.heading, self.paper.document.get('heading', '')))

        if k == 'heading_center_size':
            return (self.heading is not None
                    and self.heading.alignment == WD_ALIGN_PARAGRAPH.CENTER
                    and _all_runs(self.heading, lambda r: r.font.size is not None
                                  and abs(r.font.size.pt - check['val']) < 0.5)
                    and self._text_ok(self.heading, self.paper.document.get('heading', '')))

        if k == 'heading_bold_color':
            vals = [v.upper() for v in check['val']]
            return (self.heading is not None
                    and _all_runs(self.heading, lambda r: r.bold)
                    and _all_runs(self.heading, lambda r: (r.font.color is not None
                                                           and r.font.color.rgb is not None
                                                           and str(r.font.color.rgb).upper() in vals))
                    and self._text_ok(self.heading, self.paper.document.get('heading', '')))

        if k == 'header_text':
            for s in self.doc.sections:
                for p in s.header.paragraphs:
                    if check['val'] in p.text:
                        return True
            return False

        if k == 'footer_text':
            for s in self.doc.sections:
                for p in s.footer.paragraphs:
                    if check['val'] in p.text:
                        return True
            return False

        if k == 'footer_pagenum':
            for s in self.doc.sections:
                xml = s.footer._element.xml
                if 'PAGE' in xml and ('instrText' in xml or 'fldSimple' in xml):
                    return True
                for p in s.footer.paragraphs:
                    t = _norm(p.text)
                    if t and any(ch.isdigit() for ch in t):
                        return True
            return False

        if k == 'footer_text_bold':
            for s in self.doc.sections:
                for p in s.footer.paragraphs:
                    if check['val'] in p.text and any(
                        check['word'] in r.text and r.bold for r in _runs(p)
                    ):
                        return True
            return False

        if k == 'find_replace':
            fr = self.fr
            if fr is None:
                return False
            texts = [p.text for p in self.doc.paragraphs]
            if self.table is not None:
                for row in self.table.rows:
                    for cell in row.cells:
                        texts.append(cell.text)
            full = '\n'.join(texts)
            if fr['find'] in full:
                return False
            return full.count(fr['replace']) >= fr['min_count']

        if k == 'orientation':
            return self.sec.orientation == WD_ORIENT.LANDSCAPE

        if k == 'margins':
            s = self.sec
            return all(abs(getattr(s, a).inches - check['val']) < 0.06
                       for a in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'))

        if k == 'orient_margins_narrow':
            return (self.sec.orientation == WD_ORIENT.LANDSCAPE
                    and all(abs(getattr(self.sec, a).inches - 0.5) < 0.06
                            for a in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin')))

        if k == 'page_color':
            bg = self.doc.element.find(qn('w:background'))
            if bg is None:
                return False
            c = (bg.get(qn('w:color')) or '').upper()
            return c not in ('', 'FFFFFF', 'AUTO')

        if k == 'page_break_before':
            return self._para_task(
                check['para'],
                lambda p: p.paragraph_format.page_break_before is True
            )

        if k == 'table_borders':
            return self._table_bordered()

        if k == 'table_head_bold_shade':
            return self._head_row(bold=True, shade=True)

        if k == 'table_head_bold_center':
            return self._head_row(bold=True, center=True)

        if k == 'para_shading':
            def shaded(p):
                ppr = p._p.get_or_add_pPr()
                shd = ppr.find(qn('w:shd'))
                if shd is None:
                    return False
                return (shd.get(qn('w:fill')) or '').upper() not in ('', 'AUTO', 'FFFFFF')
            return self._para_task(check['para'], shaded)

        if k == 'page_border':
            pgb = self.sec._sectPr.find(qn('w:pgBorders'))
            if pgb is None:
                return False
            edges = [e for e in ('top', 'left', 'bottom', 'right')
                     if pgb.find(qn('w:' + e)) is not None
                     and pgb.find(qn('w:' + e)).get(qn('w:val')) not in (None, 'none', 'nil')]
            return len(edges) >= 3

        if k == 'borders_and_cell':
            if not self._table_bordered() or self.table is None:
                return False
            try:
                return _norm(self.table.rows[check['row']].cells[check['col']].text) == check['val']
            except IndexError:
                return False

        if k == 'cell_text':
            if self.table is None:
                return False
            try:
                return _norm(self.table.rows[check['row']].cells[check['col']].text) == check['val']
            except IndexError:
                return False

        if k == 'phrase_italic':
            return self._para_task(
                check['para'],
                lambda p: any(check['val'] in r.text and r.italic for r in _runs(p))
            )

        if k == 'word_double_underline':
            return self._para_task(
                check['para'],
                lambda p: any(check['val'] in r.text and r.font.underline == WD_UNDERLINE.DOUBLE
                              for r in _runs(p))
            )

        raise ValueError(f"Unknown check kind: '{k}'")

    # ── Table helpers ─────────────────────────────────────────────────────────

    def _tbl_borders_el(self):
        if self.table is None:
            return None
        tbl = self.table._tbl
        tblpr = tbl.tblPr
        return tblpr.find(qn('w:tblBorders')) if tblpr is not None else None

    def _table_bordered(self) -> bool:
        if self.table is None:
            return False
        b = self._tbl_borders_el()
        if b is not None:
            edges = [e for e in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
                     if b.find(qn('w:' + e)) is not None
                     and b.find(qn('w:' + e)).get(qn('w:val')) not in (None, 'none', 'nil')]
            if len(edges) >= 4:
                return True
        for row in self.table.rows:
            for cell in row.cells:
                tcpr = cell._tc.get_or_add_tcPr()
                cb = tcpr.find(qn('w:tcBorders'))
                if cb is None:
                    return False
                for e in ('top', 'left', 'bottom', 'right'):
                    el = cb.find(qn('w:' + e))
                    if el is None or el.get(qn('w:val')) in (None, 'none', 'nil'):
                        return False
        return True

    def _head_row(self, bold=False, shade=False, center=False) -> bool:
        if self.table is None:
            return False
        cells = self.table.rows[0].cells
        for cell in cells:
            for p in cell.paragraphs:
                if not _runs(p):
                    continue
                if bold and not _all_runs(p, lambda r: r.bold):
                    return False
                if center and p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    return False
            if shade:
                tcpr = cell._tc.get_or_add_tcPr()
                shd = tcpr.find(qn('w:shd'))
                if shd is None:
                    return False
                fill = (shd.get(qn('w:fill')) or '').upper()
                if fill in ('', 'AUTO', 'FFFFFF'):
                    return False
        return True

    # ── Main grade method ─────────────────────────────────────────────────────

    def grade(self) -> GradeResult:
        task_results: list[TaskResult] = []
        for task in self.paper.tasks:
            try:
                passed = bool(self._check(task.check))
            except Exception:
                passed = False
            earned = task.marks if passed else 0
            hint = '' if passed else task.hint_on_fail
            task_results.append(TaskResult(task=task, passed=passed, earned=earned, hint=hint))

        earned_total = sum(tr.earned for tr in task_results)
        return GradeResult(
            paper=self.paper,
            total_marks=self.paper.max_marks,
            earned_marks=earned_total,
            task_results=task_results,
        )


# ── Module-level convenience function ─────────────────────────────────────────

def grade(docx_source: str | io.BytesIO, paper: Paper) -> GradeResult:
    """Grade *docx_source* against *paper* and return a GradeResult."""
    return Grader(docx_source, paper).grade()
