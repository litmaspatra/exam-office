# ExamOffice - RSSB Word Efficiency Trainer (Desktop)

Practice the RSSB Clerk Grade-II / Junior Assistant Phase-II **Word efficiency paper
(Paper 3)** in **real MS Word** - not a browser imitation. The app opens a fresh
exam .docx in Word, shows the 15-task question paper with a 10-minute countdown,
and auto-marks your saved file out of 25 (qualifying: 9, SC/ST 7.75).

## Setup (Windows, one time)

1. Install Python from https://www.python.org/downloads/ - tick **"Add python.exe to PATH"** during install.
2. Double-click **setup.bat** (installs the one dependency, python-docx).
3. MS Word must be installed (any recent version; the exam uses Word 2013).

## Run

Double-click **run.bat**. Choose **Word Paper 1** or **Word Paper 2**, then:

- **Practice** - Word opens with the paper. Do tasks in Word, press **Ctrl+S**;
  the app ticks each task green as it detects it. Submit anytime for full marks detail.
- **Exam (10 min)** - 10:00 countdown, no live ticks, marked at the end like the hall.

In Word: do ONLY what each task asks (extra edits forfeit that task), no
Cut/Copy/Paste, right-click, Select-All or F1-F12 - same rules as the hall
(they cannot be blocked by any app, so practise the discipline yourself).

**Submit & Mark** reads your last save, so always **Ctrl+S first**.

Result screen shows marks /25, QUALIFIED / NOT QUALIFIED, and a per-task report.
Attempts are saved under My Results on the home screen.

## The papers (15 tasks, 25 marks each)

Covering every category in the board's notice:

| Category | Paper 1 (जल संरक्षण) | Paper 2 (डिजिटल साक्षरता) |
|---|---|---|
| Paragraph | justify, first-line indent | center, left indent |
| Spacing | line spacing 1.5, space after 12pt | line spacing 2.0 |
| Font | bold, size 16, red, phrase italic | bold+blue heading, italic, double underline, highlight |
| Header/Footer | header text, page number | header text, footer text + bold word |
| Paging | landscape + narrow margins | page break before |
| Background | paragraph shading | page color, page border |
| Table | head-row bold+shade, borders+cell | head-row bold+center, cell |
| Word replace | पानी -> जल | कम्प्यूटर -> कंप्यूटर |

Auto-marking reads the saved .docx directly (python-docx), so what Word shows
is exactly what gets marked. Still taught in the app: changing the given text
beyond a task forfeits that task's marks, matching the hall rule.

Note: documents use Unicode Hindi (Nirmala UI/Mangal). The hall paper is typed
in DevLys 010; the formatting tasks are identical, and typing practice belongs
to the speed tests (Papers 1-2) in the web app.
