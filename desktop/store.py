# -*- coding: utf-8 -*-
r"""
store.py – Persistent storage for ExamOffice attempt records.

All writable data lives in %LOCALAPPDATA%\ExamOffice\
  attempts\<attempt-id>\   -- per-attempt directory with document snapshot + result JSON
  papers\                  -- user-added paper folders
  examoffice.log           -- optional application log

No files are ever written inside the source tree, beside an exe, or in
PyInstaller's temp extraction directory.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import time
import uuid
from dataclasses import asdict, dataclass
from typing import Any

logger = logging.getLogger(__name__)


def resource_path(relative: str) -> str:
    """Absolute path to a bundled read-only resource; frozen-safe."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative)


# ── User-writable root ────────────────────────────────────────────────────────

def user_data_dir() -> str:
    r"""Root ExamOffice data directory: %LOCALAPPDATA%\ExamOffice or ~/.examoffice."""
    local_app_data = os.environ.get('LOCALAPPDATA', '')
    if local_app_data:
        base = os.path.join(local_app_data, 'ExamOffice')
    else:
        base = os.path.join(os.path.expanduser('~'), '.examoffice')
    os.makedirs(base, exist_ok=True)
    return base


def attempts_dir() -> str:
    d = os.path.join(user_data_dir(), 'attempts')
    os.makedirs(d, exist_ok=True)
    return d


def user_papers_dir() -> str:
    d = os.path.join(user_data_dir(), 'papers')
    os.makedirs(d, exist_ok=True)
    return d


# ── Attempt workspace ─────────────────────────────────────────────────────────

def new_attempt_dir(paper_id: str) -> tuple[str, str]:
    """
    Create a fresh per-attempt directory.
    Returns (attempt_id, attempt_dir).
    """
    attempt_id = f"{paper_id}_{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    d = os.path.join(attempts_dir(), attempt_id)
    os.makedirs(d, exist_ok=True)
    return attempt_id, d


def working_docx_path(attempt_dir: str, paper_id: str) -> str:
    """Path where the live exam document is kept."""
    return os.path.join(attempt_dir, f'ExamOffice_{paper_id}.docx')


def result_json_path(attempt_dir: str) -> str:
    return os.path.join(attempt_dir, 'result.json')


# ── Result persistence ────────────────────────────────────────────────────────

def save_result(attempt_dir: str, record: dict) -> None:
    """Persist a result record as JSON inside the attempt directory."""
    path = result_json_path(attempt_dir)
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
    except OSError as e:
        logger.warning("Could not save result to %s: %s", path, e)


def load_result(attempt_dir: str) -> dict | None:
    """Load a result record from the attempt directory, or None."""
    path = result_json_path(attempt_dir)
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


# ── History index ─────────────────────────────────────────────────────────────

def _history_path() -> str:
    return os.path.join(user_data_dir(), 'history.json')


def load_history() -> list[dict]:
    """Return list of summary records, newest first."""
    try:
        with open(_history_path(), 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, list):
                return data
    except (OSError, json.JSONDecodeError):
        pass
    return []


def append_history(record: dict) -> None:
    """Add one summary record to the history index."""
    history = load_history()
    history.insert(0, record)
    try:
        with open(_history_path(), 'w', encoding='utf-8') as f:
            json.dump(history[:500], f, ensure_ascii=False, indent=1)
    except OSError as e:
        logger.warning("Could not update history: %s", e)


def clear_history() -> None:
    try:
        os.remove(_history_path())
    except OSError:
        pass


# ── Legacy compat for old code that called store.paper_path() ─────────────────

def paper_path(paper_id: str) -> str:
    """Kept for backward-compat. Use new_attempt_dir() for new attempts."""
    d = os.path.join(user_data_dir(), 'workspace')
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f'ExamOffice_{paper_id}.docx')


# ── Logging ───────────────────────────────────────────────────────────────────

def setup_logging() -> None:
    log_path = os.path.join(user_data_dir(), 'examoffice.log')
    logging.basicConfig(
        level=logging.WARNING,
        format='%(asctime)s %(levelname)s %(name)s: %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler(sys.stderr),
        ],
    )
