# Exam Office

RSSB Hindi/English typing test simulator (new pattern), single-file app.

- Open `index.html` in any browser (double-click) - no server needed.
- 3 Hindi tests (DevLys-010 style, Remington layout) + 2 English tests (QWERTY), 10-minute countdown.
- Scoring per the new RSSB pattern: marks = correct words x 0.0625 (Hindi) / 0.05 (English), max 25, no negative marking; qualifying 9.00 (Gen/OBC) / 7.75 (SC/ST) per paper.
- Soni-style result report with QUALIFIED badge, net/gross WPM, word-by-word table, and saved history (My Results).

---

## Desktop App (Word Efficiency Trainer)

Windows desktop companion for the RSSB Clerk Grade-II / Junior Assistant Phase-II **Word Efficiency Test (Paper 3)** that tests you inside **real Microsoft Word**.

- **What it is:** Launches real Microsoft Word with a pre-loaded 15-task exam `.docx`, runs an official 10-minute countdown, and auto-marks your saved document out of 25 (MS Word must be installed).
- **Download:** Download `ExamOffice.exe` directly from the [GitHub Releases](https://github.com/litmaspatra/exam-office/releases) page (tag `v1.0.0`).
- **Run from source:** Run `desktop\setup.bat` once to install dependencies, then double-click `desktop\run.bat` to launch.
- **Build the EXE yourself:** Run `desktop\build.bat` (creates a clean virtual environment, installs dependencies + PyInstaller, and builds `desktop\dist\ExamOffice.exe`).
- **Modes:**
  - *Practice mode:* Auto-detects your saves (`Ctrl+S`) in real-time and ticks tasks green live as you complete them.
  - *Exam mode:* Strict 10-minute exam simulation without live hints; automatically evaluates your final document upon submission or time-up.
