# -*- coding: utf-8 -*-
"""ExamOffice - RSSB Clerk Grade-II / Junior Assistant Phase-II
Word efficiency trainer that uses REAL MS Word.

Flow: pick a paper (Practice or Exam) -> the app writes a fresh exam .docx and
opens it in MS Word -> you do the 15 tasks in Word -> Ctrl+S to save (Practice
mode ticks tasks on every save) -> Submit & Mark -> instant marks out of 25
with per-task report. 10 minute countdown like the exam hall.
"""
import io
import os
import subprocess
import sys
import time
import tkinter as tk
from tkinter import messagebox, ttk

# Ensure desktop directory is on sys.path regardless of CWD or frozen state
_desktop_dir = os.path.dirname(os.path.abspath(__file__))
if _desktop_dir not in sys.path:
    sys.path.insert(0, _desktop_dir)

import docx_gen
import marker
import papers as papers_mod
import store

TIME_LIMIT = 600  # 10 minutes, as per the RSSB pattern
PASS_MARKS = 9.0
PASS_MARKS_SCST = 7.75

RULES_TEXT = (
    "Exam-hall rules (RSSB notice):\n"
    "• Cut-Copy-Paste, right-click, Select-All and F1-F12 are OFF in the real exam - do not use them here either.\n"
    "• Ctrl+Z and Backspace are allowed.\n"
    "• Do ONLY what each task asks. Extra edits, or changing the given text/table, forfeit that task's marks.\n"
    "• Press Ctrl+S in Word to save - the app reads your saved file.\n"
    "• Practice mode ticks tasks green every time you save. Exam mode marks only at the end."
)


def set_dpi_awareness():
    """Enable high-DPI scaling on Windows so the UI renders crisp."""
    if os.name == 'nt':
        try:
            import ctypes
            # Per-monitor DPI aware (Windows 8.1+)
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                import ctypes
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def find_word_executable():
    """Find WINWORD.EXE path if available on Windows."""
    if os.name != 'nt':
        return None
    try:
        import winreg
        for root_key in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root_key, r"Software\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe") as key:
                    val, _ = winreg.QueryValueEx(key, "")
                    if val and os.path.exists(val):
                        return val
            except OSError:
                pass
    except Exception:
        pass

    program_files = [os.environ.get('ProgramFiles', r'C:\Program Files'),
                     os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)')]
    for pf in program_files:
        if not pf:
            continue
        candidates = [
            os.path.join(pf, 'Microsoft Office', 'root', 'Office16', 'WINWORD.EXE'),
            os.path.join(pf, 'Microsoft Office', 'Office16', 'WINWORD.EXE'),
            os.path.join(pf, 'Microsoft Office', 'root', 'Office15', 'WINWORD.EXE'),
            os.path.join(pf, 'Microsoft Office', 'Office15', 'WINWORD.EXE'),
            os.path.join(pf, 'Microsoft Office', 'Office14', 'WINWORD.EXE'),
            os.path.join(pf, 'Microsoft Office', 'Office12', 'WINWORD.EXE'),
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
    return None


def open_in_word(path):
    """Launch Microsoft Word with the specified document, raising RuntimeError if Word cannot be run."""
    if os.name == 'nt':
        word_exe = find_word_executable()
        if word_exe:
            try:
                subprocess.Popen([word_exe, path])
                return
            except Exception:
                pass
        try:
            os.startfile(path)  # noqa: S606 - intentional on Windows fallback
        except OSError as e:
            raise RuntimeError(
                "Microsoft Word is required for ExamOffice, but could not be launched.\n\n"
                "Please verify Microsoft Word is installed and associated with .docx files."
            ) from e
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def safe_grade(docx_path, paper, max_retries=6, delay=0.15):
    """Safely grade a docx file that might be currently open/locked in Word.

    Reads file bytes into an in-memory stream to release file handles immediately,
    and retries with short delays if Word is mid-save.
    """
    last_err = None
    for _ in range(max_retries):
        try:
            with open(docx_path, 'rb') as f:
                content = io.BytesIO(f.read())
            return marker.Grader(content, paper).grade()
        except Exception as e:
            last_err = e
            time.sleep(delay)
    raise last_err


def bind_mousewheel_to_canvas(canvas, widget):
    """Bind mouse wheel scrolling to canvas only while hovering over widget."""
    def _on_wheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _enter(event):
        widget.bind_all("<MouseWheel>", _on_wheel)

    def _leave(event):
        widget.unbind_all("<MouseWheel>")

    widget.bind("<Enter>", _enter)
    widget.bind("<Leave>", _leave)


class ExamOfficeApp(tk.Tk):
    def __init__(self):
        set_dpi_awareness()
        super().__init__()
        self.title('ExamOffice - RSSB Word Efficiency Trainer')
        self.geometry('1040x720')
        self.minsize(960, 640)
        self.configure(bg='#16181d')

        # Set app icon if available
        icon_path = store.resource_path('icon.ico')
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        self.session = None       # dict with live paper state
        self.timer_job = None
        self.poll_job = None

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.container = tk.Frame(self, bg='#16181d')
        self.container.pack(fill='both', expand=True)
        self.frames = {}
        for F in (HomeFrame, ExamFrame, ResultFrame):
            frame = F(self.container, self)
            self.frames[F.__name__] = frame
            frame.place(relwidth=1, relheight=1)
        self.show('HomeFrame')

    def show(self, name):
        self.frames[name].tkraise()
        if hasattr(self.frames[name], 'on_show'):
            self.frames[name].on_show()

    def _on_close(self):
        if self.session and not self.session.get('ended', False):
            if not messagebox.askyesno(
                    'Quit ExamOffice',
                    'A session is currently in progress.\n\nAre you sure you want to exit?'):
                return
        if self.timer_job:
            try:
                self.after_cancel(self.timer_job)
            except Exception:
                pass
        if self.poll_job:
            try:
                self.after_cancel(self.poll_job)
            except Exception:
                pass
        self.destroy()

    # ---- session control ----
    def start_paper(self, paper, mode):
        path = store.paper_path(paper['id'])
        try:
            docx_gen.generate(paper, path)
        except PermissionError:
            messagebox.showerror(
                'File is open in Word',
                'The paper file is currently open in Word. Close it in Word first, then press Start again.\n\n' + path)
            return
        except Exception as e:
            messagebox.showerror('Could not generate paper', str(e))
            return

        self.session = {
            'paper': paper, 'mode': mode, 'path': path,
            'deadline': time.time() + TIME_LIMIT,
            'last_mtime': None, 'ended': False, 'ticks': [False] * len(paper['tasks']),
        }

        try:
            open_in_word(path)
        except Exception as e:
            messagebox.showerror('Could Not Open Microsoft Word', str(e))
            self.session = None
            return

        self.show('ExamFrame')
        self._tick_timer()
        self._poll_file()

    def _tick_timer(self):
        s = self.session
        if not s or s['ended']:
            return
        left = int(s['deadline'] - time.time())
        if left <= 0:
            self.frames['ExamFrame'].set_timer('0:00', low=True)
            s['ended'] = True
            messagebox.showinfo(
                'Time up',
                'Time is over. Press Ctrl+S in Word to save your work, then click OK to get your marks.')
            self.submit(auto=True)
            return
        self.frames['ExamFrame'].set_timer('%d:%02d' % divmod(left, 60), low=left <= 60)
        self.timer_job = self.after(500, self._tick_timer)

    def _poll_file(self):
        s = self.session
        if not s or s['ended']:
            return
        try:
            mt = os.path.getmtime(s['path'])
            if mt != s['last_mtime']:
                if s['mode'] == 'practice':
                    try:
                        r = safe_grade(s['path'], s['paper'])
                        s['ticks'] = r['per_task']
                        self.frames['ExamFrame'].update_ticks(r['per_task'])
                        s['last_mtime'] = mt
                        self.frames['ExamFrame'].set_status('Saved in Word - ' + time.strftime('%H:%M:%S'))
                    except Exception:
                        # File is mid-save; leave last_mtime unchanged so next poll cycle retries
                        pass
                else:
                    s['last_mtime'] = mt
                    self.frames['ExamFrame'].set_status('Saved in Word - ' + time.strftime('%H:%M:%S'))
        except OSError:
            pass
        self.poll_job = self.after(1000, self._poll_file)

    def submit(self, auto=False):
        s = self.session
        if not s:
            return
        if not auto:
            if not messagebox.askyesno(
                    'Submit & Mark',
                    'Did you press Ctrl+S in Word? Your marks come from the saved file.\n\nMark now?'):
                return
        s['ended'] = True
        if self.timer_job:
            try:
                self.after_cancel(self.timer_job)
            except Exception:
                pass
        if self.poll_job:
            try:
                self.after_cancel(self.poll_job)
            except Exception:
                pass

        try:
            r = safe_grade(s['path'], s['paper'])
        except Exception as e:
            messagebox.showerror('Could not read the file',
                                 'Could not read your saved file:\n' + str(e) +
                                 '\n\nMake sure the document is saved and not corrupted.')
            self.show('HomeFrame')
            return

        time_used = min(TIME_LIMIT, max(1, int(time.time() - (s['deadline'] - TIME_LIMIT))))
        store.save_attempt({
            'paper': s['paper']['title'], 'mode': s['mode'],
            'marks': r['marks'], 'tasks': r['tasks_done'], 'time_used': time_used,
        })
        self.frames['ResultFrame'].set_result(s['paper'], s['mode'], r, time_used)
        self.session = None
        self.show('ResultFrame')

    def quit_paper(self):
        s = self.session
        if s and not s['ended']:
            if not messagebox.askyesno('Quit paper', 'Quit without marks? Your saved file stays in the workspace folder.'):
                return
            s['ended'] = True
            if self.timer_job:
                try:
                    self.after_cancel(self.timer_job)
                except Exception:
                    pass
            if self.poll_job:
                try:
                    self.after_cancel(self.poll_job)
                except Exception:
                    pass
        self.session = None
        self.show('HomeFrame')


# ---------- UI components ----------

def _btn(parent, text, cmd, accent=False, width=None):
    b = tk.Button(parent, text=text, command=cmd, font=('Segoe UI', 10, 'bold' if accent else 'normal'),
                  bg='#e8650e' if accent else '#262a33', fg='white',
                  activebackground='#f0761c' if accent else '#31363f',
                  activeforeground='white', relief='flat', padx=18, pady=8, cursor='hand2',
                  width=width if width else 0)
    return b


class HomeFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg='#16181d')
        self.app = app
        tk.Label(self, text='ExamOffice', font=('Segoe UI', 24, 'bold'), bg='#16181d', fg='#e8650e').pack(pady=(20, 2))
        tk.Label(self, text='RSSB Clerk Grade-II / Junior Assistant Phase-II - Word Efficiency (Paper 3) with REAL MS Word',
                 font=('Segoe UI', 10), bg='#16181d', fg='#9aa0aa').pack(pady=(0, 14))

        cards = tk.Frame(self, bg='#16181d')
        cards.pack()
        self.card_frames = []
        for paper in papers_mod.PAPERS:
            card = tk.Frame(cards, bg='#1e222b', padx=18, pady=14)
            card.pack(side='left', padx=12, pady=6)
            tk.Label(card, text=paper['title'], font=('Segoe UI', 12, 'bold'), bg='#1e222b', fg='white').pack(anchor='w')
            tk.Label(card, text='15 tasks  ·  25 marks  ·  10 min  ·  min 9 (SC/ST 7.75)',
                     font=('Segoe UI', 9), bg='#1e222b', fg='#9aa0aa').pack(anchor='w', pady=(2, 8))
            best = tk.Label(card, text='', font=('Segoe UI', 9), bg='#1e222b', fg='#f0a35e')
            best.pack(anchor='w')
            row = tk.Frame(card, bg='#1e222b')
            row.pack(pady=(8, 0))
            _btn(row, 'Practice', lambda p=paper: app.start_paper(p, 'practice')).pack(side='left', padx=(0, 8))
            _btn(row, 'Exam (10 min)', lambda p=paper: app.start_paper(p, 'exam'), accent=True).pack(side='left')
            self.card_frames.append((paper['id'], best))

        hist_box = tk.LabelFrame(self, text=' My Results ', font=('Segoe UI', 10), bg='#16181d', fg='#9aa0aa',
                                 bd=1, labelanchor='n')
        hist_box.pack(fill='both', expand=True, padx=30, pady=14)

        hist_inner = tk.Frame(hist_box, bg='#12141a')
        hist_inner.pack(fill='both', expand=True, padx=8, pady=8)

        hist_scroll = ttk.Scrollbar(hist_inner, orient='vertical')
        self.hist_list = tk.Listbox(hist_inner, font=('Consolas', 10), bg='#12141a', fg='#d6d9de',
                                    selectbackground='#2b2f38', relief='flat', activestyle='none',
                                    yscrollcommand=hist_scroll.set)
        hist_scroll.config(command=self.hist_list.yview)
        self.hist_list.pack(side='left', fill='both', expand=True)
        hist_scroll.pack(side='right', fill='y')

        _btn(hist_box, 'Clear history', self._clear).pack(pady=(0, 8))

    def _clear(self):
        store.clear_history()
        self.on_show()

    def on_show(self):
        hist = store.load_history()
        self.hist_list.delete(0, 'end')
        for e in hist[:100]:
            self.hist_list.insert('end', '%-16s  %-34s  %-8s  %5.2f / 25  (%d/15 tasks, %d:%02d)' % (
                e.get('ts', ''), e.get('paper', ''), e.get('mode', ''),
                e.get('marks', 0), e.get('tasks', 0), *divmod(e.get('time_used', 0), 60)))
        best = {}
        for e in hist:
            k = e.get('paper', '')
            best[k] = max(best.get(k, 0), e.get('marks', 0))
        for pid, lbl in self.card_frames:
            title = next(p['title'] for p in papers_mod.PAPERS if p['id'] == pid)
            lbl.config(text=('Best: %.2f / 25' % best[title]) if title in best else 'Not attempted yet')


class ExamFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg='#16181d')
        self.app = app

        top = tk.Frame(self, bg='#16181d')
        top.pack(fill='x', padx=20, pady=(12, 4))
        self.title_lbl = tk.Label(top, text='', font=('Segoe UI', 13, 'bold'), bg='#16181d', fg='white')
        self.title_lbl.pack(side='left')
        self.timer_lbl = tk.Label(top, text='10:00', font=('Consolas', 22, 'bold'), bg='#16181d', fg='#7dd3fc')
        self.timer_lbl.pack(side='left', padx=30)
        self.mode_lbl = tk.Label(top, text='', font=('Segoe UI', 10), bg='#16181d', fg='#9aa0aa')
        self.mode_lbl.pack(side='left')

        mid = tk.Frame(self, bg='#16181d')
        mid.pack(fill='both', expand=True, padx=20, pady=4)

        rules = tk.Label(mid, text=RULES_TEXT, font=('Segoe UI', 9), bg='#1e222b', fg='#d6d9de',
                         justify='left', anchor='nw', padx=12, pady=8)
        rules.pack(side='top', fill='x', pady=(0, 6))

        list_box = tk.Frame(mid, bg='#12141a')
        list_box.pack(fill='both', expand=True)
        self.canvas = tk.Canvas(list_box, bg='#12141a', highlightthickness=0)
        scroll = ttk.Scrollbar(list_box, orient='vertical', command=self.canvas.yview)
        self.task_frame = tk.Frame(self.canvas, bg='#12141a')
        self.task_frame_window = self.canvas.create_window((0, 0), window=self.task_frame, anchor='nw')

        self.task_frame.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        bind_mousewheel_to_canvas(self.canvas, self.canvas)
        bind_mousewheel_to_canvas(self.canvas, self.task_frame)

        self.task_labels = []

        bot = tk.Frame(self, bg='#16181d')
        bot.pack(fill='x', padx=20, pady=10)
        _btn(bot, 'Open paper in Word again', self._reopen).pack(side='left')
        self.status_lbl = tk.Label(bot, text='', font=('Segoe UI', 9), bg='#16181d', fg='#9aa0aa')
        self.status_lbl.pack(side='left', padx=16)
        _btn(bot, 'Quit', app.quit_paper).pack(side='right', padx=(8, 0))
        _btn(bot, 'Submit & Mark', app.submit, accent=True).pack(side='right')

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.task_frame_window, width=event.width)
        wrap = max(500, event.width - 60)
        for lbl, txt in self.task_labels:
            txt.config(wraplength=wrap)

    def _reopen(self):
        s = self.app.session
        if s:
            try:
                open_in_word(s['path'])
            except Exception as e:
                messagebox.showerror('Could Not Open Word', str(e))

    def on_show(self):
        s = self.app.session
        if not s:
            return
        self.title_lbl.config(text=s['paper']['title'])
        self.mode_lbl.config(text=('Practice mode - tasks tick on every Ctrl+S' if s['mode'] == 'practice'
                                   else 'Exam mode - marked only at the end'))
        for w in self.task_frame.winfo_children():
            w.destroy()
        self.task_labels = []
        wrap = max(500, self.canvas.winfo_width() - 60)
        for i, t in enumerate(s['paper']['tasks']):
            lbl = tk.Label(self.task_frame, text='☐', font=('Segoe UI', 10), bg='#12141a', fg='#9aa0aa', anchor='w')
            lbl.grid(row=i, column=0, sticky='nw', padx=(10, 6), pady=3)
            txt = tk.Label(self.task_frame, text='%d. %s   [%d mark%s]' % (i + 1, t['t'], t['m'], 's' if t['m'] > 1 else ''),
                           font=('Segoe UI', 10), bg='#12141a', fg='#d6d9de', anchor='w', justify='left', wraplength=wrap)
            txt.grid(row=i, column=1, sticky='w', pady=3)
            bind_mousewheel_to_canvas(self.canvas, txt)
            bind_mousewheel_to_canvas(self.canvas, lbl)
            self.task_labels.append((lbl, txt))
        self.set_status('Paper opened in MS Word. Do the tasks there, Ctrl+S to save.')

    def update_ticks(self, per):
        for (lbl, txt), ok in zip(self.task_labels, per):
            lbl.config(text='✅' if ok else '☐')
            txt.config(fg='#4ade80' if ok else '#d6d9de')

    def set_timer(self, text, low=False):
        self.timer_lbl.config(text=text, fg='#f87171' if low else '#7dd3fc')

    def set_status(self, text):
        self.status_lbl.config(text=text)


class ResultFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg='#16181d')
        self.app = app
        self.marks_lbl = tk.Label(self, text='', font=('Segoe UI', 32, 'bold'), bg='#16181d', fg='white')
        self.marks_lbl.pack(pady=(20, 0))
        self.badge_lbl = tk.Label(self, text='', font=('Segoe UI', 13, 'bold'), bg='#16181d')
        self.badge_lbl.pack(pady=4)
        self.sub_lbl = tk.Label(self, text='', font=('Segoe UI', 10), bg='#16181d', fg='#9aa0aa')
        self.sub_lbl.pack(pady=(0, 8))

        box = tk.Frame(self, bg='#12141a')
        box.pack(fill='both', expand=True, padx=40, pady=4)
        self.canvas = tk.Canvas(box, bg='#12141a', highlightthickness=0)
        scroll = ttk.Scrollbar(box, orient='vertical', command=self.canvas.yview)
        self.rows = tk.Frame(self.canvas, bg='#12141a')
        self.rows_window = self.canvas.create_window((0, 0), window=self.rows, anchor='nw')

        self.rows.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        bind_mousewheel_to_canvas(self.canvas, self.canvas)
        bind_mousewheel_to_canvas(self.canvas, self.rows)

        bot = tk.Frame(self, bg='#16181d')
        bot.pack(pady=10)
        _btn(bot, 'Back to Home', lambda: app.show('HomeFrame')).pack(side='left', padx=8)
        self.retry_btn = _btn(bot, 'Retry this paper', self._retry, accent=True)
        self.retry_btn.pack(side='left', padx=8)
        self._last = None
        self._row_labels = []

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.rows_window, width=event.width)
        wrap = max(450, event.width - 90)
        for txt in self._row_labels:
            txt.config(wraplength=wrap)

    def _retry(self):
        if self._last:
            self.app.start_paper(*self._last)

    def set_result(self, paper, mode, r, time_used):
        self._last = (paper, mode)
        self.marks_lbl.config(text='%.2f / 25' % r['marks'])
        if r['marks'] >= PASS_MARKS:
            self.badge_lbl.config(
                text='QUALIFIED  (needed 9.00 Gen/OBC, 7.75 SC/ST) - cleared by %.2f' % (r['marks'] - PASS_MARKS),
                fg='#4ade80'
            )
        elif r['marks'] >= PASS_MARKS_SCST:
            self.badge_lbl.config(
                text='QUALIFIED only for SC/ST (7.75) - need %.2f more for Gen/OBC' % (PASS_MARKS - r['marks']),
                fg='#fbbf24'
            )
        else:
            self.badge_lbl.config(
                text='NOT QUALIFIED - need %.2f more marks' % (PASS_MARKS - r['marks']),
                fg='#f87171'
            )
        self.sub_lbl.config(text='%s  ·  %s mode  ·  %d of 15 tasks correct  ·  time used %d:%02d'
                            % (paper['title'], mode, r['tasks_done'], *divmod(time_used, 60)))
        for w in self.rows.winfo_children():
            w.destroy()
        self._row_labels = []
        wrap = max(450, self.canvas.winfo_width() - 90)
        for i, (t, ok) in enumerate(zip(paper['tasks'], r['per_task'])):
            lbl = tk.Label(self.rows, text='✓ +%d' % t['m'] if ok else '✗ 0', font=('Consolas', 10, 'bold'),
                           bg='#12141a', fg='#4ade80' if ok else '#f87171', width=6, anchor='e')
            lbl.grid(row=i, column=0, padx=(10, 8), pady=3, sticky='ne')
            txt = tk.Label(self.rows, text='%d. %s' % (i + 1, t['t']), font=('Segoe UI', 10),
                           bg='#12141a', fg='#d6d9de', anchor='w', justify='left', wraplength=wrap)
            txt.grid(row=i, column=1, sticky='w', pady=3)
            bind_mousewheel_to_canvas(self.canvas, lbl)
            bind_mousewheel_to_canvas(self.canvas, txt)
            self._row_labels.append(txt)


def main():
    app = ExamOfficeApp()
    app.mainloop()


if __name__ == '__main__':
    main()
