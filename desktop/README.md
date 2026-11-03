# ExamOffice Desktop

Python 3.10+ desktop app using Tkinter and python-docx.

## Requirements

- Windows 10 / 11
- Microsoft Word 2013+ (must be installed and associated with .docx files)
- Python 3.10+

## Setup

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python exam_office.py
```

Or run `setup.bat` then `run.bat`.

## Build

```bat
build.bat
```

Produces `dist\ExamOffice.exe` (single file, no installer needed).

## App Data Location

All runtime data is written to `%LOCALAPPDATA%\ExamOffice\`:

```
%LOCALAPPDATA%\ExamOffice\
  attempts\           ← per-attempt directories with saved .docx and result JSON
  papers\             ← user-added paper folders
  history.json        ← result history index
  examoffice.log      ← application log
```

## Papers

Built-in papers are in `papers\w-1\` and `papers\w-2\`.

To add a custom paper, create a folder in `%LOCALAPPDATA%\ExamOffice\papers\<your-id>\`
containing a `paper.json` file. See [PAPERS.md](PAPERS.md) for the full schema.

## Modules

| File | Purpose |
|---|---|
| `exam_office.py` | Main entry point, Tkinter UI |
| `paper_loader.py` | Load and validate paper.json manifests |
| `docx_gen.py` | Generate exam .docx from a Paper object |
| `marker.py` | Auto-mark a saved .docx against paper checks |
| `store.py` | Attempt directories, result persistence, logging |

## Testing

```bat
..\.venv\Scripts\python.exe -m pytest tests\ -v
```
