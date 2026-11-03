# ExamOffice — Word Efficiency Trainer

A free, offline Windows desktop app for practising RSSB computer efficiency tests.
Launches real Microsoft Word with a preloaded exam document, tracks your progress live,
and marks your saved work automatically.

**No internet required. No account. No subscription.**

---

## Download

Get the latest installer-free executable from [Releases](https://github.com/litmaspatra/exam-office/releases/latest).

---

## Requirements

- Windows 10 / 11
- Microsoft Word (2013 or later, any edition — 365 / 2021 / 2019 / 2016 / 2013)
- No installation needed; run `ExamOffice.exe` directly

---

## Quick Start

1. Download `ExamOffice.exe` from the Releases page.
2. Double-click to launch.
3. Select a paper (**Word Paper 1** or **Word Paper 2**).
4. Choose **Practice** (live checkmarks as you save) or **Exam** (marks calculated upon test completion).
5. Complete tasks in Microsoft Word, pressing **Ctrl+S** after each one.
6. Click **Finish** or wait for the timer to expire — your score appears instantly.

---

## Run from Source

```bat
git clone https://github.com/litmaspatra/exam-office.git
cd exam-office
python -m venv .venv
.venv\Scripts\activate
pip install -r desktop\requirements.txt
python desktop\exam_office.py
```

Requires Python 3.10+.

---

## Build the Executable

```bat
cd desktop
build.bat
```

Output: `desktop\dist\ExamOffice.exe`

---

## Papers

Two built-in papers (15 tasks, 25 marks each):

| Paper | Topic |
|---|---|
| Word Paper 1 | जल संरक्षण |
| Word Paper 2 | डिजिटल साक्षरता |

To add your own paper, see [desktop/PAPERS.md](desktop/PAPERS.md).

---

## Scoring

- **Pass**: ≥ 9.0 marks
- **Pass (SC/ST)**: ≥ 7.75 marks
- Marks shown per task with hints on failure.

---

## App Data

All attempt files, logs, and history are stored in:
```
%LOCALAPPDATA%\ExamOffice\
```

---

## License

MIT License — see [LICENSE](LICENSE)
