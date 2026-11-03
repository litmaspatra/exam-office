# -*- coding: utf-8 -*-
"""
exam_office.py – ExamOffice desktop application entry point.

Presents a start screen with paper selection, opens a fresh exam .docx in the
user's installed Microsoft Word, manages the countdown, evaluates the saved
document, and displays a per-task result breakdown.

Architecture notes:
  • Timer uses deadline arithmetic (time.monotonic), not tick counting, so it
    remains accurate even if Tkinter callbacks are delayed.
  • All Tk widget updates run on the Tk main thread; no Tk calls from threads.
  • Finalization is guarded by a single flag so an attempt cannot be marked twice.
  • Word integration uses WINWORD.EXE directly when found, falling back to
    os.startfile; missing Word is detected and reported with a dialog before
    the attempt starts.
  • Per-attempt directories isolate documents so retries cannot overwrite
    earlier work.
"""

from __future__ import annotations

import ctypes
import io
import logging
import os
import subprocess
import sys
import time
import tkinter as tk
import winreg
from tkinter import messagebox, ttk

# Ensure desktop/ is on sys.path regardless of CWD (works frozen and as script)
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import docx_gen
import marker
import store
from paper_loader import Paper, load_all_papers, get_user_papers_dir

logger = logging.getLogger(__name__)

PASS_MARKS = 9.0
PASS_MARKS_SCST = 7.75

# Colour palette
C_BG = '#16181d'
C_CARD = '#1e222b'
C_CARD2 = '#12141a'
C_ACCENT = '#e8650e'
C_ACCENT2 = '#f0761c'
C_TEXT = '#d6d9de'
C_DIM = '#9aa0aa'
C_GREEN = '#4ade80'
C_RED = '#f87171'
C_AMBER = '#fbbf24'
C_BLUE = '#7dd3fc'

FONT_TITLE = ('Segoe UI', 20, 'bold')
FONT_BODY = ('Segoe UI', 10)
FONT_BODY_B = ('Segoe UI', 10, 'bold')
FONT_SMALL = ('Segoe UI', 9)
FONT_MONO = ('Consolas', 11, 'bold')
FONT_BIG = ('Segoe UI', 30, 'bold')

RULES_TEXT = (
    "Exam-hall rules (RSSB notice):\n"
    "• Do ONLY what each task asks — extra edits forfeit marks for that task.\n"
    "• Cut/Copy/Paste, right-click, Select-All and F1–F12 are NOT permitted.\n"
    "• Ctrl+Z (Undo) and Backspace are allowed.\n"
    "• Press Ctrl+S in Word after completing each task.\n"
    "• Marking uses the LAST SAVED state — unsaved edits are invisible to this app."
)


# ── DPI awareness ─────────────────────────────────────────────────────────────

def _set_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ── Word detection ────────────────────────────────────────────────────────────

def _find_winword() -> str | None:
    """Return full path to WINWORD.EXE, or None."""
    if os.name != 'nt':
        return None
    try:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    hive,
                    r"Software\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
                ) as key:
                    val, _ = winreg.QueryValueEx(key, "")
                    if val and os.path.isfile(val):
                        return val
            except OSError:
                pass
    except Exception:
        pass
    for pf in [os.environ.get('ProgramFiles', ''),
               os.environ.get('ProgramFiles(x86)', '')]:
        if not pf:
            continue
        for sub in ('root\\Office16', 'Office16', 'root\\Office15', 'Office15',
                    'Office14', 'Office12'):
            candidate = os.path.join(pf, 'Microsoft Office', sub, 'WINWORD.EXE')
            if os.path.isfile(candidate):
                return candidate
    return None


def _open_in_word(path: str) -> None:
    """
    Open *path* in Microsoft Word.
    Raises RuntimeError with a user-friendly message if Word cannot be launched.
    """
    word_exe = _find_winword()
    if word_exe:
        try:
            subprocess.Popen([word_exe, path])
            return
        except Exception as e:
            logger.error("WINWORD.EXE launch failed: %s", e)
    if os.name == 'nt':
        try:
            os.startfile(path)  # noqa: S606
            return
        except OSError as e:
            raise RuntimeError(
                "Microsoft Word could not be launched.\n\n"
                "ExamOffice requires Microsoft Word to be installed and associated with .docx files.\n\n"
                f"System error: {e}"
            ) from e
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


# ── Safe grading with retries (handles Word file locks) ───────────────────────

def _safe_grade(docx_path: str, paper: Paper,
                max_retries: int = 8,
                delay: float = 0.15) -> marker.GradeResult | None:
    """
    Read docx_path into memory then grade it.
    Retries to tolerate transient Word file locks and incomplete writes.
    Returns None if all attempts fail (caller should treat as ungraded, not zero).
    """
    last_err: Exception | None = None
    lock_name = os.path.join(os.path.dirname(docx_path),
                             '~$' + os.path.basename(docx_path))
    for attempt in range(max_retries):
        if attempt > 0:
            time.sleep(delay)
        # Skip if Word's temporary lock file is present (mid-save)
        if os.path.exists(lock_name) and attempt < max_retries - 1:
            continue
        try:
            with open(docx_path, 'rb') as f:
                buf = io.BytesIO(f.read())
            return marker.Grader(buf, paper).grade()
        except Exception as e:
            last_err = e
    logger.error("Could not grade %s after %d retries: %s", docx_path, max_retries, last_err)
    return None


# ── Widget helpers ────────────────────────────────────────────────────────────

def _btn(parent, text: str, cmd, accent=False, small=False, **kwargs) -> tk.Button:
    fs = 9 if small else 10
    weight = 'bold' if accent else 'normal'
    return tk.Button(
        parent, text=text, command=cmd,
        font=('Segoe UI', fs, weight),
        bg=C_ACCENT if accent else '#262a33',
        fg='white',
        activebackground=C_ACCENT2 if accent else '#31363f',
        activeforeground='white',
        relief='flat', padx=14, pady=6 if small else 8,
        cursor='hand2',
        **kwargs
    )


def _label(parent, text='', font=None, fg=C_TEXT, bg=C_BG, **kwargs) -> tk.Label:
    return tk.Label(parent, text=text, font=font or FONT_BODY,
                    fg=fg, bg=bg, **kwargs)


def _scrollable_frame(parent, bg=C_CARD2) -> tuple[tk.Canvas, tk.Frame, ttk.Scrollbar]:
    """Return (canvas, inner_frame, scrollbar). Caller must pack canvas + scrollbar."""
    canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
    vsb = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    frame = tk.Frame(canvas, bg=bg)
    win = canvas.create_window((0, 0), window=frame, anchor='nw')

    def _on_frame_configure(e):
        canvas.configure(scrollregion=canvas.bbox('all'))

    def _on_canvas_configure(e):
        canvas.itemconfig(win, width=e.width)

    frame.bind('<Configure>', _on_frame_configure)
    canvas.bind('<Configure>', _on_canvas_configure)

    def _bind_wheel(widget):
        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        widget.bind('<Enter>', lambda e: widget.bind_all('<MouseWheel>', _wheel))
        widget.bind('<Leave>', lambda e: widget.unbind_all('<MouseWheel>'))

    _bind_wheel(canvas)
    _bind_wheel(frame)
    return canvas, frame, vsb


# ── Application ───────────────────────────────────────────────────────────────

class ExamOfficeApp(tk.Tk):
    def __init__(self) -> None:
        _set_dpi_awareness()
        super().__init__()

        store.setup_logging()
        self.title('ExamOffice – Word Efficiency Trainer')
        self.geometry('1060x720')
        self.minsize(940, 620)
        self.configure(bg=C_BG)

        icon = store.resource_path('icon.ico')
        if os.path.isfile(icon):
            try:
                self.iconbitmap(icon)
            except Exception:
                pass

        self.protocol('WM_DELETE_WINDOW', self._on_close)

        # Live session state
        self._session: dict | None = None      # set during an active attempt
        self._timer_job: str | None = None
        self._poll_job: str | None = None
        self._finalized: bool = False           # guard against double-marking

        # Load papers
        self._papers, self._paper_errors = load_all_papers()

        # Build frame container
        self._container = tk.Frame(self, bg=C_BG)
        self._container.pack(fill='both', expand=True)
        self._frames: dict[str, tk.Frame] = {}
        for F in (HomeFrame, AttemptFrame, ResultFrame):
            name = F.__name__
            frame = F(self._container, self)
            self._frames[name] = frame
            frame.place(relwidth=1, relheight=1)

        self._show('HomeFrame')

    # ── Navigation ──────────────────────────────────────────────────────────

    def _show(self, name: str) -> None:
        f = self._frames[name]
        f.tkraise()
        if hasattr(f, 'on_show'):
            f.on_show()

    # ── Session control ──────────────────────────────────────────────────────

    def start_attempt(self, paper: Paper, mode: str) -> None:
        """Begin a new attempt. mode is 'practice' or 'exam'."""
        # Check Word is available BEFORE creating files
        if os.name == 'nt' and _find_winword() is None:
            messagebox.showerror(
                'Microsoft Word Not Found',
                "ExamOffice could not find Microsoft Word on this computer.\n\n"
                "Please install Microsoft Word (any recent version) and try again.\n\n"
                "ExamOffice requires real Microsoft Word — it cannot work with\n"
                "LibreOffice, WPS, or other editors."
            )
            return

        attempt_id, attempt_dir = store.new_attempt_dir(paper.id)
        docx_path = store.working_docx_path(attempt_dir, paper.id)

        try:
            docx_gen.generate(paper, docx_path)
        except PermissionError:
            messagebox.showerror(
                'File Busy',
                f"The exam document is open elsewhere. Please close it and try again.\n\n{docx_path}"
            )
            return
        except Exception as e:
            messagebox.showerror('Document Error', f"Could not create the exam document:\n\n{e}")
            return

        try:
            _open_in_word(docx_path)
        except RuntimeError as e:
            messagebox.showerror('Could Not Open Word', str(e))
            return

        self._session = {
            'paper': paper,
            'mode': mode,
            'attempt_id': attempt_id,
            'attempt_dir': attempt_dir,
            'docx_path': docx_path,
            'deadline': time.monotonic() + paper.duration_seconds,
            'start_wall': time.time(),
            'last_mtime': None,
            'ticks': [False] * len(paper.tasks),
        }
        self._finalized = False

        self._show('AttemptFrame')
        self._tick_timer()
        self._poll_file()

    def _tick_timer(self) -> None:
        s = self._session
        if s is None or self._finalized:
            return
        remaining = s['deadline'] - time.monotonic()
        if remaining <= 0:
            self._frames['AttemptFrame'].set_timer('0:00', low=True)
            self._finalize(reason='time_up')
            return
        mins, secs = divmod(int(remaining), 60)
        self._frames['AttemptFrame'].set_timer(f'{mins}:{secs:02d}', low=remaining <= 60)
        self._timer_job = self.after(500, self._tick_timer)

    def _poll_file(self) -> None:
        s = self._session
        if s is None or self._finalized:
            return
        try:
            mt = os.path.getmtime(s['docx_path'])
            if mt != s['last_mtime']:
                if s['mode'] == 'practice':
                    result = _safe_grade(s['docx_path'], s['paper'])
                    if result is not None:
                        s['ticks'] = [tr.passed for tr in result.task_results]
                        self._frames['AttemptFrame'].update_ticks(s['ticks'])
                        s['last_mtime'] = mt
                        ts = time.strftime('%H:%M:%S')
                        done = sum(s['ticks'])
                        self._frames['AttemptFrame'].set_status(
                            f'Saved {ts} — {done}/{len(s["ticks"])} tasks detected'
                        )
                    # If grade failed, keep last_mtime unchanged → retry on next poll
                else:
                    s['last_mtime'] = mt
                    self._frames['AttemptFrame'].set_status(
                        f"Saved at {time.strftime('%H:%M:%S')}"
                    )
        except OSError:
            pass
        self._poll_job = self.after(1000, self._poll_file)

    def request_finish(self) -> None:
        """Called by the Finish button."""
        if self._finalized:
            return
        if not messagebox.askyesno(
            'Finish & Mark',
            'Press Ctrl+S in Word first to save any last edits.\n\n'
            'Mark this attempt now?'
        ):
            return
        self._finalize(reason='user_finish')

    def _cancel_jobs(self) -> None:
        for attr in ('_timer_job', '_poll_job'):
            job = getattr(self, attr)
            if job is not None:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                setattr(self, attr, None)

    def _finalize(self, reason: str) -> None:
        if self._finalized:
            return
        self._finalized = True
        self._cancel_jobs()

        s = self._session
        if s is None:
            return

        if reason == 'time_up':
            messagebox.showinfo(
                'Time Up',
                'The 10 minutes have ended.\n\n'
                'Press Ctrl+S in Word NOW to save your final work, then click OK.'
            )

        # Grade the saved document
        result = _safe_grade(s['docx_path'], s['paper'])

        time_used = max(1, int(time.monotonic() - (s['deadline'] - s['paper'].duration_seconds)))
        time_used = min(s['paper'].duration_seconds, time_used)

        if result is None:
            # Ungraded — could not read the file
            record = {
                'attempt_id': s['attempt_id'],
                'paper_id': s['paper'].id,
                'paper_title': s['paper'].title,
                'mode': s['mode'],
                'reason': reason,
                'ungraded': True,
                'ungraded_reason': 'Could not read the saved document.',
                'earned_marks': 0,
                'total_marks': s['paper'].max_marks,
                'time_used': time_used,
                'timestamp': time.strftime('%Y-%m-%d %H:%M'),
            }
            store.save_result(s['attempt_dir'], record)
            store.append_history(record)
            self._session = None
            self._frames['ResultFrame'].set_ungraded(s['paper'], s['mode'], reason, time_used)
            self._show('ResultFrame')
            return

        # Build persistence record
        task_records = []
        for tr in result.task_results:
            task_records.append({
                'task_id': tr.task.id,
                'task_text': tr.task.text,
                'marks': tr.task.marks,
                'earned': tr.earned,
                'passed': tr.passed,
                'hint': tr.hint,
            })
        record = {
            'attempt_id': s['attempt_id'],
            'paper_id': s['paper'].id,
            'paper_title': s['paper'].title,
            'mode': s['mode'],
            'reason': reason,
            'ungraded': False,
            'earned_marks': result.earned_marks,
            'total_marks': result.total_marks,
            'tasks': task_records,
            'time_used': time_used,
            'timestamp': time.strftime('%Y-%m-%d %H:%M'),
        }
        store.save_result(s['attempt_dir'], record)
        store.append_history(record)
        self._session = None
        self._frames['ResultFrame'].set_result(s['paper'], s['mode'], result, reason, time_used)
        self._show('ResultFrame')

    def abort_attempt(self) -> None:
        s = self._session
        if s is not None and not self._finalized:
            if not messagebox.askyesno(
                'Abandon Attempt',
                'Are you sure you want to abandon this attempt?\n\n'
                'No result will be saved. Your Word document remains in the attempt folder.'
            ):
                return
        self._cancel_jobs()
        self._session = None
        self._finalized = True
        self._show('HomeFrame')

    def reload_papers(self) -> None:
        self._papers, self._paper_errors = load_all_papers()
        self._frames['HomeFrame'].on_show()

    # ── Window close ─────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        if self._session and not self._finalized:
            if not messagebox.askyesno(
                'Exit ExamOffice',
                'An attempt is in progress.\n\nExit anyway?'
            ):
                return
        self._cancel_jobs()
        self.destroy()


# ── Home Screen ───────────────────────────────────────────────────────────────

class HomeFrame(tk.Frame):
    def __init__(self, parent, app: ExamOfficeApp) -> None:
        super().__init__(parent, bg=C_BG)
        self._app = app
        self._paper_rows: list[tuple[str, tk.Label]] = []   # (paper_id, best_label)
        self._build()

    def _build(self) -> None:
        # Header
        hdr = tk.Frame(self, bg=C_BG)
        hdr.pack(fill='x', padx=30, pady=(20, 4))
        _label(hdr, 'ExamOffice', font=FONT_TITLE, fg=C_ACCENT, bg=C_BG).pack(side='left')
        _btn(hdr, '⟳ Reload', self._app.reload_papers, small=True).pack(side='right', padx=4)
        _btn(hdr, '📂 Papers folder', self._open_papers_folder, small=True).pack(side='right', padx=4)

        _label(self,
               'RSSB Clerk Grade-II / Junior Assistant Phase-II — Word Efficiency Test (Paper 3)',
               font=FONT_SMALL, fg=C_DIM, bg=C_BG).pack()

        # Prerequisite note
        note = tk.Frame(self, bg='#1a1c23', padx=16, pady=10)
        note.pack(fill='x', padx=30, pady=(10, 4))
        _label(note, '📋  Prerequisite: Microsoft Word must be installed on this PC.',
               font=FONT_SMALL, fg='#f0a35e', bg='#1a1c23', anchor='w').pack(fill='x')
        _label(note,
               'This app opens the exam document in real MS Word and marks your saved file. '
               'It cannot see unsaved edits.',
               font=FONT_SMALL, fg=C_DIM, bg='#1a1c23', anchor='w').pack(fill='x')

        # Paper cards area (scrollable for many papers)
        cards_outer = tk.Frame(self, bg=C_BG)
        cards_outer.pack(fill='x', padx=30, pady=6)
        self._cards_frame = tk.Frame(cards_outer, bg=C_BG)
        self._cards_frame.pack()

        # History
        hist_frame = tk.LabelFrame(self, text='  My Results  ',
                                   font=FONT_SMALL, bg=C_BG, fg=C_DIM, bd=1)
        hist_frame.pack(fill='both', expand=True, padx=30, pady=(6, 12))

        hist_inner = tk.Frame(hist_frame, bg=C_CARD2)
        hist_inner.pack(fill='both', expand=True, padx=6, pady=6)
        canvas, self._hist_frame, vsb = _scrollable_frame(hist_inner, bg=C_CARD2)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        btn_row = tk.Frame(hist_frame, bg=C_BG)
        btn_row.pack(pady=(0, 6))
        _btn(btn_row, 'Clear History', self._clear_history, small=True).pack()

    def _rebuild_cards(self) -> None:
        for w in self._cards_frame.winfo_children():
            w.destroy()
        self._paper_rows = []

        if not self._app._papers:
            msg = 'No papers found.'
            if self._app._paper_errors:
                msg += '\n\nErrors:\n' + '\n'.join(self._app._paper_errors[:3])
            _label(self._cards_frame, msg, font=FONT_BODY, fg=C_RED, bg=C_BG).pack(pady=10)
            return

        for paper in self._app._papers:
            card = tk.Frame(self._cards_frame, bg=C_CARD, padx=20, pady=14)
            card.pack(side='left', padx=10, pady=4, anchor='n')

            _label(card, paper.title, font=FONT_BODY_B, fg='white', bg=C_CARD).pack(anchor='w')
            mins = paper.duration_seconds // 60
            info = f'{len(paper.tasks)} tasks  ·  {paper.max_marks} marks  ·  {mins} min'
            _label(card, info, font=FONT_SMALL, fg=C_DIM, bg=C_CARD).pack(anchor='w', pady=(2, 0))
            pass_info = f'Pass: {paper.pass_marks} (SC/ST {paper.pass_marks_scst})'
            _label(card, pass_info, font=FONT_SMALL, fg=C_DIM, bg=C_CARD).pack(anchor='w')

            best_lbl = _label(card, 'Not attempted yet', font=FONT_SMALL,
                              fg='#f0a35e', bg=C_CARD)
            best_lbl.pack(anchor='w', pady=(4, 8))
            self._paper_rows.append((paper.id, best_lbl))

            btn_row = tk.Frame(card, bg=C_CARD)
            btn_row.pack(anchor='w')
            _btn(btn_row, 'Practice',
                 lambda p=paper: self._app.start_attempt(p, 'practice')).pack(side='left', padx=(0, 8))
            _btn(btn_row, 'Exam',
                 lambda p=paper: self._app.start_attempt(p, 'exam'),
                 accent=True).pack(side='left')

    def _rebuild_history(self) -> None:
        for w in self._hist_frame.winfo_children():
            w.destroy()

        history = store.load_history()
        if not history:
            _label(self._hist_frame, 'No attempts yet.', font=FONT_SMALL,
                   fg=C_DIM, bg=C_CARD2).grid(row=0, column=0, columnspan=4, padx=10, pady=6)
            return

        headers = ['Date', 'Paper', 'Mode', 'Score', 'Tasks', 'Time']
        for c, h in enumerate(headers):
            _label(self._hist_frame, h, font=FONT_BODY_B, fg=C_DIM, bg=C_CARD2).grid(
                row=0, column=c, padx=(10, 16), pady=(4, 2), sticky='w')

        for i, rec in enumerate(history[:100], start=1):
            ts = rec.get('timestamp', '')
            title = rec.get('paper_title', '')[:28]
            mode = rec.get('mode', '')
            ungraded = rec.get('ungraded', False)
            score_text = '—' if ungraded else f"{rec.get('earned_marks', 0):.1f} / {rec.get('total_marks', 25)}"
            tasks_text = '—' if ungraded else f"{sum(1 for t in rec.get('tasks', []) if t.get('passed'))}/{len(rec.get('tasks', []))}"
            tu = rec.get('time_used', 0)
            time_text = f'{tu // 60}:{tu % 60:02d}'
            row_fg = C_RED if ungraded else C_TEXT

            for c, val in enumerate([ts, title, mode, score_text, tasks_text, time_text]):
                _label(self._hist_frame, val, font=FONT_SMALL, fg=row_fg, bg=C_CARD2).grid(
                    row=i, column=c, padx=(10, 16), pady=1, sticky='w')

        # Update best scores on cards
        best: dict[str, float] = {}
        for rec in history:
            if not rec.get('ungraded'):
                k = rec.get('paper_id', '')
                best[k] = max(best.get(k, 0.0), float(rec.get('earned_marks', 0)))
        for pid, lbl in self._paper_rows:
            if pid in best:
                lbl.config(text=f'Best: {best[pid]:.1f} / 25')

    def on_show(self) -> None:
        self._rebuild_cards()
        self._rebuild_history()

        # Show any paper load errors
        if self._app._paper_errors:
            err_text = 'Paper load warnings:\n' + '\n'.join(self._app._paper_errors)
            _label(self, err_text, font=FONT_SMALL, fg=C_AMBER, bg=C_BG,
                   justify='left').pack(padx=30, pady=2, anchor='w')

    def _clear_history(self) -> None:
        if messagebox.askyesno('Clear History', 'Delete all result history?'):
            store.clear_history()
            self._rebuild_history()

    def _open_papers_folder(self) -> None:
        folder = get_user_papers_dir()
        os.makedirs(folder, exist_ok=True)
        if os.name == 'nt':
            os.startfile(folder)  # noqa: S606


# ── Attempt Screen ────────────────────────────────────────────────────────────

class AttemptFrame(tk.Frame):
    def __init__(self, parent, app: ExamOfficeApp) -> None:
        super().__init__(parent, bg=C_BG)
        self._app = app
        self._task_labels: list[tuple[tk.Label, tk.Label]] = []
        self._build()

    def _build(self) -> None:
        # Top bar: paper name | timer | mode
        top = tk.Frame(self, bg=C_BG)
        top.pack(fill='x', padx=20, pady=(12, 4))
        self._title_lbl = _label(top, '', font=('Segoe UI', 12, 'bold'), fg='white', bg=C_BG)
        self._title_lbl.pack(side='left')
        self._timer_lbl = tk.Label(top, text='10:00', font=FONT_MONO,
                                   fg=C_BLUE, bg=C_BG)
        self._timer_lbl.pack(side='left', padx=28)
        self._mode_lbl = _label(top, '', font=FONT_SMALL, fg=C_DIM, bg=C_BG)
        self._mode_lbl.pack(side='left')

        # Rules banner
        rules = _label(self, RULES_TEXT, font=FONT_SMALL, fg=C_TEXT,
                       bg='#1a1c23', justify='left', anchor='nw', padx=14, pady=8)
        rules.pack(fill='x', padx=20, pady=(0, 4))

        # Document path
        path_frame = tk.Frame(self, bg=C_BG)
        path_frame.pack(fill='x', padx=20, pady=(0, 4))
        _label(path_frame, 'Document:', font=FONT_SMALL, fg=C_DIM, bg=C_BG).pack(side='left')
        self._path_lbl = _label(path_frame, '', font=FONT_SMALL, fg='#93c5fd', bg=C_BG)
        self._path_lbl.pack(side='left', padx=6)
        _btn(path_frame, 'Open in Word', self._reopen_word, small=True).pack(side='left', padx=4)

        # Task list
        list_outer = tk.Frame(self, bg=C_CARD2)
        list_outer.pack(fill='both', expand=True, padx=20, pady=4)
        canvas, self._task_inner, vsb = _scrollable_frame(list_outer, bg=C_CARD2)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self._canvas = canvas

        # Bottom bar
        bot = tk.Frame(self, bg=C_BG)
        bot.pack(fill='x', padx=20, pady=8)
        self._status_lbl = _label(bot, '', font=FONT_SMALL, fg=C_DIM, bg=C_BG)
        self._status_lbl.pack(side='left')
        _btn(bot, 'Abandon', self._app.abort_attempt).pack(side='right', padx=(6, 0))
        _btn(bot, 'Finish & Mark', self._app.request_finish, accent=True).pack(side='right')

    def on_show(self) -> None:
        s = self._app._session
        if s is None:
            return
        paper = s['paper']
        self._title_lbl.config(text=paper.title)
        self._mode_lbl.config(text=(
            'PRACTICE — tasks tick live on every Ctrl+S'
            if s['mode'] == 'practice'
            else 'EXAM — no live feedback; scored on Finish'
        ))
        self._path_lbl.config(text=s['docx_path'])
        self.set_timer('10:00')
        self.set_status('Document opened in MS Word — do the tasks and press Ctrl+S to save.')

        for w in self._task_inner.winfo_children():
            w.destroy()
        self._task_labels = []
        wrap = max(500, self._canvas.winfo_width() - 80)
        for i, task in enumerate(paper.tasks):
            chk = _label(self._task_inner, '☐', font=FONT_BODY, fg=C_DIM, bg=C_CARD2)
            chk.grid(row=i, column=0, padx=(10, 6), pady=3, sticky='nw')
            txt = _label(
                self._task_inner,
                f'{i + 1}. {task.text}   [{task.marks} mark{"s" if task.marks > 1 else ""}]',
                font=FONT_BODY, fg=C_TEXT, bg=C_CARD2,
                anchor='w', justify='left', wraplength=wrap
            )
            txt.grid(row=i, column=1, pady=3, sticky='w')
            self._task_labels.append((chk, txt))

    def _reopen_word(self) -> None:
        s = self._app._session
        if s:
            try:
                _open_in_word(s['docx_path'])
            except RuntimeError as e:
                messagebox.showerror('Could Not Open Word', str(e))

    def update_ticks(self, ticks: list[bool]) -> None:
        for (chk, txt), ok in zip(self._task_labels, ticks):
            chk.config(text='✅' if ok else '☐')
            txt.config(fg=C_GREEN if ok else C_TEXT)

    def set_timer(self, text: str, low: bool = False) -> None:
        self._timer_lbl.config(text=text, fg=C_RED if low else C_BLUE)

    def set_status(self, text: str) -> None:
        self._status_lbl.config(text=text)


# ── Result Screen ─────────────────────────────────────────────────────────────

class ResultFrame(tk.Frame):
    def __init__(self, parent, app: ExamOfficeApp) -> None:
        super().__init__(parent, bg=C_BG)
        self._app = app
        self._retry_args: tuple | None = None
        self._row_labels: list[tk.Label] = []
        self._build()

    def _build(self) -> None:
        # Score area
        score_frame = tk.Frame(self, bg=C_BG)
        score_frame.pack(pady=(18, 0))
        self._score_lbl = _label(score_frame, '', font=FONT_BIG, fg='white', bg=C_BG)
        self._score_lbl.pack()
        self._badge_lbl = _label(score_frame, '', font=('Segoe UI', 13, 'bold'), bg=C_BG)
        self._badge_lbl.pack(pady=2)
        self._sub_lbl = _label(score_frame, '', font=FONT_SMALL, fg=C_DIM, bg=C_BG)
        self._sub_lbl.pack()

        # Per-task breakdown
        detail_outer = tk.Frame(self, bg=C_CARD2)
        detail_outer.pack(fill='both', expand=True, padx=30, pady=(10, 4))
        canvas, self._detail_inner, vsb = _scrollable_frame(detail_outer, bg=C_CARD2)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self._canvas = canvas

        # Bottom buttons
        bot = tk.Frame(self, bg=C_BG)
        bot.pack(pady=10)
        _btn(bot, '← Back to Papers', lambda: self._app._show('HomeFrame')).pack(side='left', padx=8)
        self._retry_btn = _btn(bot, 'Retry', self._retry, accent=True)
        self._retry_btn.pack(side='left', padx=8)

    def set_ungraded(self, paper: Paper, mode: str, reason: str, time_used: int) -> None:
        self._retry_args = (paper, mode)
        self._score_lbl.config(text='— / 25', fg=C_DIM)
        self._badge_lbl.config(text='UNGRADED', fg=C_AMBER)
        self._sub_lbl.config(
            text=f'{paper.title}  ·  {mode}  ·  {reason}  ·  '
                 f'Time: {time_used // 60}:{time_used % 60:02d}'
        )
        for w in self._detail_inner.winfo_children():
            w.destroy()
        _label(
            self._detail_inner,
            'The saved document could not be read.\n\n'
            'Possible causes:\n'
            '  • Word had not finished saving when time expired.\n'
            '  • The file was deleted or moved.\n'
            '  • The .docx was corrupted.\n\n'
            'Your document file is preserved in the attempt folder.',
            font=FONT_BODY, fg=C_AMBER, bg=C_CARD2, justify='left'
        ).grid(row=0, column=0, padx=16, pady=12, sticky='w')

    def set_result(self, paper: Paper, mode: str, result: marker.GradeResult,
                   reason: str, time_used: int) -> None:
        self._retry_args = (paper, mode)
        earned = result.earned_marks
        total = result.total_marks

        self._score_lbl.config(text=f'{earned:.1f} / {total}', fg='white')

        if earned >= paper.pass_marks:
            self._badge_lbl.config(
                text=f'QUALIFIED  (needed {paper.pass_marks}  Gen/OBC, {paper.pass_marks_scst}  SC/ST) — '
                     f'cleared by {earned - paper.pass_marks:.2f}',
                fg=C_GREEN
            )
        elif earned >= paper.pass_marks_scst:
            self._badge_lbl.config(
                text=f'QUALIFIED for SC/ST only — need {paper.pass_marks - earned:.2f} more for Gen/OBC',
                fg=C_AMBER
            )
        else:
            self._badge_lbl.config(
                text=f'NOT QUALIFIED — need {paper.pass_marks - earned:.2f} more marks',
                fg=C_RED
            )

        reason_text = {'user_finish': 'Submitted', 'time_up': 'Time expired'}.get(reason, reason)
        tasks_done = sum(1 for tr in result.task_results if tr.passed)
        self._sub_lbl.config(
            text=(f'{paper.title}  ·  {mode}  ·  {reason_text}  ·  '
                  f'{tasks_done}/{len(result.task_results)} tasks  ·  '
                  f'Time: {time_used // 60}:{time_used % 60:02d}')
        )

        for w in self._detail_inner.winfo_children():
            w.destroy()
        self._row_labels = []
        wrap = max(450, self._canvas.winfo_width() - 100)

        # Column headers
        _label(self._detail_inner, 'Score', font=FONT_BODY_B, fg=C_DIM, bg=C_CARD2
               ).grid(row=0, column=0, padx=(12, 8), pady=(6, 2), sticky='e')
        _label(self._detail_inner, 'Task', font=FONT_BODY_B, fg=C_DIM, bg=C_CARD2
               ).grid(row=0, column=1, padx=(0, 8), pady=(6, 2), sticky='w')
        _label(self._detail_inner, 'Note', font=FONT_BODY_B, fg=C_DIM, bg=C_CARD2
               ).grid(row=0, column=2, padx=(0, 8), pady=(6, 2), sticky='w')

        for i, tr in enumerate(result.task_results, start=1):
            score_text = f'✓ +{tr.earned}' if tr.passed else f'✗  0/{tr.task.marks}'
            score_fg = C_GREEN if tr.passed else C_RED
            note_text = '' if tr.passed else (tr.hint or 'Task not completed correctly.')

            _label(self._detail_inner, score_text, font=('Consolas', 10, 'bold'),
                   fg=score_fg, bg=C_CARD2, width=8, anchor='e'
                   ).grid(row=i, column=0, padx=(12, 8), pady=2, sticky='ne')
            txt = _label(self._detail_inner,
                         f'{i}. {tr.task.text}',
                         font=FONT_SMALL, fg=C_TEXT, bg=C_CARD2,
                         anchor='w', justify='left', wraplength=wrap)
            txt.grid(row=i, column=1, pady=2, sticky='w', padx=(0, 8))
            note_lbl = _label(self._detail_inner, note_text,
                              font=FONT_SMALL, fg=C_AMBER if not tr.passed else C_DIM,
                              bg=C_CARD2, anchor='w', justify='left', wraplength=max(200, wrap // 2))
            note_lbl.grid(row=i, column=2, pady=2, sticky='w', padx=(0, 12))
            self._row_labels.extend([txt, note_lbl])

    def _retry(self) -> None:
        if self._retry_args:
            self._app.start_attempt(*self._retry_args)


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    app = ExamOfficeApp()
    app.mainloop()


if __name__ == '__main__':
    main()
