# -*- coding: utf-8 -*-
"""App data: workspace paths + attempt history (no Tk dependency)."""
import json, os, sys, time

def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def workspace_dir():
    d = os.path.join(app_dir(), 'workspace')
    os.makedirs(d, exist_ok=True)
    return d

def paper_path(paper_id):
    return os.path.join(workspace_dir(), 'ExamOffice_%s.docx' % paper_id)

def history_path():
    return os.path.join(app_dir(), 'examoffice_history.json')

def load_history():
    try:
        with open(history_path(), 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []

def save_attempt(entry):
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
    try:
        os.remove(history_path())
    except OSError:
        pass
