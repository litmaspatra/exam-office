# -*- coding: utf-8 -*-
r"""
paper_loader.py – Load and validate ExamOffice paper manifests (paper.json).

Built-in papers live alongside this module in papers/.
User-added papers are loaded from %LOCALAPPDATA%\ExamOffice\papers\  (each in its own subfolder).

Never executes Python, eval(), shell commands, or arbitrary imports from a paper file.
All asset paths are resolved relative to the paper folder; traversal outside is rejected.
"""

from __future__ import annotations

import json
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Any

# ── Supported check kinds ──────────────────────────────────────────────────────
SUPPORTED_KINDS = frozenset([
    'align', 'bold', 'italic', 'underline',
    'linespacing', 'indent', 'first_line_indent', 'spaceafter',
    'color', 'highlight',
    'heading_bold', 'heading_size', 'heading_color',
    'heading_center_size', 'heading_bold_color',
    'phrase_italic', 'word_double_underline',
    'header_text', 'footer_text', 'footer_pagenum', 'footer_text_bold',
    'find_replace',
    'orientation', 'margins', 'orient_margins_narrow',
    'page_color', 'page_border', 'page_break_before',
    'para_shading',
    'table_borders', 'table_head_bold_shade', 'table_head_bold_center',
    'borders_and_cell', 'cell_text',
])

SCHEMA_VERSION = 1


class PaperLoadError(Exception):
    """Raised when a paper.json fails validation."""


@dataclass
class Task:
    id: str
    text: str
    marks: int
    check: dict
    hint_on_fail: str = ''


@dataclass
class Paper:
    id: str
    title: str
    duration_seconds: int
    max_marks: int
    pass_marks: float
    pass_marks_scst: float
    document: dict
    tasks: list[Task]
    folder: str  # absolute path to the paper folder

    @property
    def find_replace(self):
        return self.document.get('find_replace')


def resource_path(relative: str) -> str:
    """Return the absolute path to a bundled resource (frozen-safe)."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative)


def _validate_check(check: Any, paper_id: str, task_id: str) -> None:
    """Raise PaperLoadError if a check dict is malformed."""
    if not isinstance(check, dict):
        raise PaperLoadError(f"[{paper_id}/{task_id}] check must be a dict, got {type(check).__name__}")
    kind = check.get('kind')
    if not isinstance(kind, str) or kind not in SUPPORTED_KINDS:
        raise PaperLoadError(
            f"[{paper_id}/{task_id}] unknown check kind '{kind}'. "
            f"Supported: {sorted(SUPPORTED_KINDS)}"
        )


def _safe_path_check(folder: str, relative: str) -> str:
    """Return absolute path, rejecting any traversal outside folder."""
    abs_path = os.path.normpath(os.path.join(folder, relative))
    if not abs_path.startswith(os.path.normpath(folder)):
        raise PaperLoadError(f"Path traversal rejected: '{relative}' goes outside paper folder")
    return abs_path


def _load_from_folder(folder: str) -> Paper:
    """Load and validate a single paper from its folder."""
    manifest = os.path.join(folder, 'paper.json')
    if not os.path.isfile(manifest):
        raise PaperLoadError(f"No paper.json found in: {folder}")

    try:
        with open(manifest, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise PaperLoadError(f"Invalid JSON in {manifest}: {e}") from e

    if not isinstance(data, dict):
        raise PaperLoadError(f"paper.json root must be an object in: {folder}")

    ver = data.get('schema_version')
    if ver != SCHEMA_VERSION:
        raise PaperLoadError(f"schema_version must be {SCHEMA_VERSION}, got {ver!r} in: {folder}")

    paper_id = data.get('id', '')
    if not isinstance(paper_id, str) or not re.match(r'^[\w-]+$', paper_id):
        raise PaperLoadError(f"Invalid or missing 'id' in: {folder}")

    title = data.get('title', '')
    if not isinstance(title, str) or not title.strip():
        raise PaperLoadError(f"[{paper_id}] 'title' must be a non-empty string")

    duration = data.get('duration_seconds', 600)
    if not isinstance(duration, int) or duration < 60:
        raise PaperLoadError(f"[{paper_id}] 'duration_seconds' must be an int >= 60")

    max_marks = data.get('max_marks', 25)
    if not isinstance(max_marks, int) or max_marks < 1:
        raise PaperLoadError(f"[{paper_id}] 'max_marks' must be a positive int")

    pass_marks = float(data.get('pass_marks', 9.0))
    pass_marks_scst = float(data.get('pass_marks_scst', 7.75))

    doc = data.get('document')
    if not isinstance(doc, dict):
        raise PaperLoadError(f"[{paper_id}] 'document' must be a dict")
    if not isinstance(doc.get('heading'), str):
        raise PaperLoadError(f"[{paper_id}] document.heading must be a string")
    if not isinstance(doc.get('paragraphs'), list) or len(doc['paragraphs']) < 1:
        raise PaperLoadError(f"[{paper_id}] document.paragraphs must be a non-empty list")
    table = doc.get('table')
    if table is not None:
        if not isinstance(table, dict) or not isinstance(table.get('head'), list):
            raise PaperLoadError(f"[{paper_id}] document.table must have a 'head' list")

    raw_tasks = data.get('tasks')
    if not isinstance(raw_tasks, list) or len(raw_tasks) == 0:
        raise PaperLoadError(f"[{paper_id}] 'tasks' must be a non-empty list")

    tasks: list[Task] = []
    seen_ids: set[str] = set()
    total_marks = 0
    for i, t in enumerate(raw_tasks):
        if not isinstance(t, dict):
            raise PaperLoadError(f"[{paper_id}] task[{i}] must be a dict")
        tid = t.get('id', '')
        if not isinstance(tid, str) or not tid:
            raise PaperLoadError(f"[{paper_id}] task[{i}] missing 'id'")
        if tid in seen_ids:
            raise PaperLoadError(f"[{paper_id}] duplicate task id '{tid}'")
        seen_ids.add(tid)

        text = t.get('text', '')
        if not isinstance(text, str) or not text.strip():
            raise PaperLoadError(f"[{paper_id}/{tid}] 'text' must be non-empty")

        marks = t.get('marks', 0)
        if not isinstance(marks, int) or marks < 1:
            raise PaperLoadError(f"[{paper_id}/{tid}] 'marks' must be a positive int")

        check = t.get('check')
        _validate_check(check, paper_id, tid)

        hint = t.get('hint_on_fail', '')
        tasks.append(Task(id=tid, text=text, marks=marks, check=check, hint_on_fail=hint))
        total_marks += marks

    if total_marks != max_marks:
        raise PaperLoadError(
            f"[{paper_id}] tasks total marks ({total_marks}) != max_marks ({max_marks})"
        )

    return Paper(
        id=paper_id,
        title=title,
        duration_seconds=duration,
        max_marks=max_marks,
        pass_marks=pass_marks,
        pass_marks_scst=pass_marks_scst,
        document=doc,
        tasks=tasks,
        folder=folder,
    )


def load_all_papers(extra_dirs: list[str] | None = None) -> tuple[list[Paper], list[str]]:
    """
    Load all valid papers. Returns (papers, errors).

    Search order:
      1. Built-in papers bundled with the app (papers/ next to this module or in _MEIPASS).
      2. User-added folders in %LOCALAPPDATA%\\ExamOffice\\papers\\ (or extra_dirs).
    """
    search_dirs: list[str] = []

    # Built-in papers
    builtin_dir = resource_path('papers')
    if os.path.isdir(builtin_dir):
        search_dirs.append(builtin_dir)

    # User papers from LOCALAPPDATA
    local_app_data = os.environ.get('LOCALAPPDATA', '')
    if local_app_data:
        user_papers_dir = os.path.join(local_app_data, 'ExamOffice', 'papers')
        os.makedirs(user_papers_dir, exist_ok=True)
        search_dirs.append(user_papers_dir)

    if extra_dirs:
        search_dirs.extend(extra_dirs)

    papers: list[Paper] = []
    errors: list[str] = []
    seen_ids: set[str] = set()

    for search_dir in search_dirs:
        if not os.path.isdir(search_dir):
            continue
        for entry in sorted(os.listdir(search_dir)):
            folder = os.path.join(search_dir, entry)
            if not os.path.isdir(folder):
                continue
            if not os.path.isfile(os.path.join(folder, 'paper.json')):
                continue
            try:
                p = _load_from_folder(folder)
                if p.id in seen_ids:
                    errors.append(f"Skipping duplicate paper id '{p.id}' in: {folder}")
                    continue
                seen_ids.add(p.id)
                papers.append(p)
            except PaperLoadError as e:
                errors.append(str(e))

    return papers, errors


def get_user_papers_dir() -> str:
    local_app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    return os.path.join(local_app_data, 'ExamOffice', 'papers')
