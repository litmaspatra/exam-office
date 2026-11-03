# -*- coding: utf-8 -*-
"""ExamOffice - RSSB Clerk Grade-II / Junior Assistant Phase-II
Word efficiency trainer that uses REAL MS Word.

Flow: pick a paper (Practice or Exam) -> the app writes a fresh exam .docx and
opens it in MS Word -> you do the 15 tasks in Word -> Ctrl+S to save (Practice
mode ticks tasks on every save) -> Submit & Mark -> instant marks out of 25
with per-task report. 10 minute countdown like the exam hall.
"""
import os
import subprocess
import sys
import time
import tkinter as tk
from tkinter import messagebox, ttk

import papers as papers_mod
import docx_gen
import marker
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


def open_in_word(path):
    if os.name == 'nt':
        os.startfile(path)  # noqa: S606 - intentional on Windows
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


class ExamOfficeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ExamOffice - RSSB Word Efficiency Trainer')
        self.geometry('1020x680')
        self.minsize(940, 620)
        self.configure(bg='#16181d')

        self.session = None       # dict with live paper state
        self.timer_job = None
        self.poll_job = None

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

    # ---- session control ----
    def start_paper(self, paper, mode):
        path = store.paper_path(paper['id'])
        try:
            docx_gen.generate(paper, path)
        except PermissionError:
            messagebox.showerror(
                'File is open in Word',
                'The paper file is open in Word. Close it in Word first, then press Start again.\n\n' + path)
            return
        self.session = {
            'paper': paper, 'mode': mode, 'path': path,
            'deadline': time.time() + TIME_LIMIT,
            'last_mtime': None, 'ended': False, 'ticks': [False] * len(paper['tasks']),
        }
        try:
            open_in_word(path)
        except Exception as e:
            messagebox.showerror('Could not open Word', str(e))
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
                s['last_mtime'] = mt
                if s['mode'] == 'practice':
                    self._recheck()
                self.frames['ExamFrame'].set_status('Saved in Word - ' + time.strftime('%H:%M:%S'))
        except OSError:
            pass
        self.poll_job = self.after(1000, self._poll_file)

    def _recheck(self):
        s = self.session
        try:
            r = marker.grade(s['path'], s['paper'])
            s['ticks'] = r['per_task']
            self.frames['ExamFrame'].update_ticks(r['per_task'])
        except Exception:
            pass  # file mid-save; next poll will retry

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
            self.after_cancel(self.timer_job)
        if self.poll_job:
            self.after_cancel(self.poll_job)
        try:
            r = marker.grade(s['path'], s['paper'])
        except Exception as e:
            messagebox.showerror('Could not read the file', str(e))
            self.show('HomeFrame')
            return
        time_used = min(TIME_LIMIT, int(time.time() - (s['deadline'] - TIME_LIMIT)))
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
                self.after_cancel(self.timer_job)
            if self.poll_job:
                self.after_cancel(self.poll_job)
        self.session = None
        self.show('HomeFrame')


# ---------- frames ----------

def _btn(parent, text, cmd, accent=False, width=None):
    b = tk.Button(parent, text=text, command=cmd, font=('Segoe UI', 11, 'bold' if accent else 'normal'),
                  bg='#e8650e' if accent else '#262a33', fg='white', activebackground='#f0761c' if accent else '#31363f',
                  activeforeground='white', relief='flat', padx=18, pady=8, cursor='hand2',
                  width=width if width else 0)
    return b


class HomeFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg='#16181d')
        self.app = app
        tk.Label(self, text='ExamOffice', font=('Segoe UI', 22, 'bold'), bg='#16181d', fg='#e8650e').pack(pady=(22, 0))
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
        self.hist_list = tk.Listbox(hist_box, font=('Consolas', 10), bg='#12141a', fg='#d6d9de',
                                    selectbackground='#2b2f38', relief='flat', activestyle='none')
        self.hist_list.pack(fill='both', expand=True, padx=8, pady=8)
        _btn(hist_box, 'Clear history', self._clear).pack(pady=(0, 8))

    def _clear(self):
        store.clear_history()
        self.on_show()

    def on_show(self):
        hist = store.load_history()
        self.hist_list.delete(0, 'end')
        for e in hist[:50]:
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
        top.pack(fill='x', padx=20, pady=(14, 6))
        self.title_lbl = tk.Label(top, text='', font=('Segoe UI', 14, 'bold'), bg='#16181d', fg='white')
        self.title_lbl.pack(side='left')
        self.timer_lbl = tk.Label(top, text='10:00', font=('Consolas', 22, 'bold'), bg='#16181d', fg='#7dd3fc')
        self.timer_lbl.pack(side='left', padx=40)
        self.mode_lbl = tk.Label(top, text='', font=('Segoe UI', 10), bg='#16181d', fg='#9aa0aa')
        self.mode_lbl.pack(side='left')

        mid = tk.Frame(self, bg='#16181d')
        mid.pack(fill='both', expand=True, padx=20, pady=6)

        rules = tk.Label(mid, text=RULES_TEXT, font=('Segoe UI', 9), bg='#1e222b', fg='#d6d9de',
                         justify='left', anchor='nw', padx=12, pady=10)
        rules.pack(side='top', fill='x', pady=(0, 8))

        list_box = tk.Frame(mid, bg='#12141a')
        list_box.pack(fill='both', expand=True)
        canvas = tk.Canvas(list_box, bg='#12141a', highlightthickness=0)
        scroll = ttk.Scrollbar(list_box, orient='vertical', command=canvas.yview)
        self.task_frame = tk.Frame(canvas, bg='#12141a')
        self.task_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.task_frame, anchor='nw')
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')
        self.task_labels = []

        bot = tk.Frame(self, bg='#16181d')
        bot.pack(fill='x', padx=20, pady=10)
        _btn(bot, 'Open paper in Word again', self._reopen).pack(side='left')
        self.status_lbl = tk.Label(bot, text='', font=('Segoe UI', 9), bg='#16181d', fg='#9aa0aa')
        self.status_lbl.pack(side='left', padx=16)
        _btn(bot, 'Quit', app.quit_paper).pack(side='right', padx=(8, 0))
        _btn(bot, 'Submit & Mark', app.submit, accent=True).pack(side='right')

    def _reopen(self):
        s = self.app.session
        if s:
            try:
                open_in_word(s['path'])
            except Exception as e:
                messagebox.showerror('Could not open Word', str(e))

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
        for i, t in enumerate(s['paper']['tasks']):
            lbl = tk.Label(self.task_frame, text='☐', font=('Segoe UI', 10), bg='#12141a', fg='#9aa0aa', anchor='w')
            lbl.grid(row=i, column=0, sticky='w', padx=(10, 4), pady=2)
            txt = tk.Label(self.task_frame, text='%d. %s   [%d mark%s]' % (i + 1, t['t'], t['m'], 's' if t['m'] > 1 else ''),
                           font=('Segoe UI', 10), bg='#12141a', fg='#d6d9de', anchor='w', justify='left', wraplength=760)
            txt.grid(row=i, column=1, sticky='w', pady=2)
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
        self.marks_lbl = tk.Label(self, text='', font=('Segoe UI', 34, 'bold'), bg='#16181d', fg='white')
        self.marks_lbl.pack(pady=(26, 0))
        self.badge_lbl = tk.Label(self, text='', font=('Segoe UI', 13, 'bold'), bg='#16181d')
        self.badge_lbl.pack(pady=4)
        self.sub_lbl = tk.Label(self, text='', font=('Segoe UI', 10), bg='#16181d', fg='#9aa0aa')
        self.sub_lbl.pack(pady=(0, 10))

        box = tk.Frame(self, bg='#12141a')
        box.pack(fill='both', expand=True, padx=60, pady=6)
        canvas = tk.Canvas(box, bg='#12141a', highlightthickness=0)
        scroll = ttk.Scrollbar(box, orient='vertical', command=canvas.yview)
        self.rows = tk.Frame(canvas, bg='#12141a')
        self.rows.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.rows, anchor='nw')
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        bot = tk.Frame(self, bg='#16181d')
        bot.pack(pady=12)
        _btn(bot, 'Back to Home', lambda: app.show('HomeFrame')).pack(side='left', padx=8)
        self.retry_btn = _btn(bot, 'Retry this paper', self._retry, accent=True)
        self.retry_btn.pack(side='left', padx=8)
        self._last = None

    def _retry(self):
        if self._last:
            self.app.start_paper(*self._last)

    def set_result(self, paper, mode, r, time_used):
        self._last = (paper, mode)
        self.marks_lbl.config(text='%.2f / 25' % r['marks'])
        if r['marks'] >= PASS_MARKS:
            self.badge_lbl.config(text='QUALIFIED  (needed 9.00 Gen/OBC, 7.75 SC/ST) - cleared by %.2f' % (r['marks'] - PASS_MARKS), fg='#4ade80')
        elif r['marks'] >= PASS_MARKS_SCST:
            self.badge_lbl.config(text='QUALIFIED only for SC/ST (7.75) - need %.2f more for Gen/OBC' % (PASS_MARKS - r['marks']), fg='#fbbf24')
        else:
            self.badge_lbl.config(text='NOT QUALIFIED - need %.2f more marks' % (PASS_MARKS - r['marks']), fg='#f87171')
        self.sub_lbl.config(text='%s  ·  %s mode  ·  %d of 15 tasks correct  ·  time used %d:%02d'
                            % (paper['title'], mode, r['tasks_done'], *divmod(time_used, 60)))
        for w in self.rows.winfo_children():
            w.destroy()
        for i, (t, ok) in enumerate(zip(paper['tasks'], r['per_task'])):
            tk.Label(self.rows, text='✓ +%d' % t['m'] if ok else '✗ 0', font=('Consolas', 10, 'bold'),
                     bg='#12141a', fg='#4ade80' if ok else '#f87171', width=6, anchor='e').grid(row=i, column=0, padx=(10, 8), pady=2)
            tk.Label(self.rows, text='%d. %s' % (i + 1, t['t']), font=('Segoe UI', 10),
                     bg='#12141a', fg='#d6d9de', anchor='w', justify='left', wraplength=680).grid(row=i, column=1, sticky='w', pady=2)


def main():
    app = ExamOfficeApp()
    app.mainloop()


if __name__ == '__main__':
    main()
