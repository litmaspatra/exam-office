# -*- coding: utf-8 -*-
"""App data: user storage paths + attempt history.

All runtime data (generated documents, history) is stored in %LOCALAPPDATA%\\ExamOffice.
No files are written inside the application directory or temp extraction folder.
"""
import json
import os
import sys
import time


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def user_data_dir():
    """Get writable user directory in %LOCALAPPDATA%\\ExamOffice."""
    local_app_data = os.environ.get('LOCALAPPDATA')
    if local_app_data:
        base = os.path.join(local_app_data, 'ExamOffice')
    else:
        base = os.path.join(os.path.expanduser('~'), '.examoffice')
    os.makedirs(base, exist_ok=True)
    return base


def workspace_dir():
    """Workspace folder where active exam .docx files are saved."""
    d = os.path.join(user_data_dir(), 'workspace')
    os.makedirs(d, exist_ok=True)
    return d


def paper_path(paper_id):
    """Path to the exam document for a paper."""
    return os.path.join(workspace_dir(), 'ExamOffice_%s.docx' % paper_id)


def history_path():
    """Path to the attempt history JSON file."""
    return os.path.join(user_data_dir(), 'examoffice_history.json')


def load_history():
    """Load attempt history from user data directory."""
    try:
        with open(history_path(), 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def save_attempt(entry):
    """Save an exam attempt to user data directory."""
    h = load_history()
    entry = dict(entry)
    entry['ts'] = time.strftime('%Y-%m-%d %H:%M')
    h.insert(0, entry)
    try:
        with open(history_path(), 'w', encoding='utf-8') as f:
            json.dump(h[:200], f, ensure_ascii=False, indent=1)
    except Exception:
        pass
    return h


def clear_history():
    """Clear attempt history."""
    try:
        os.remove(history_path())
    except OSError:
        pass
