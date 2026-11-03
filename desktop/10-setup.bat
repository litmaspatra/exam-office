@echo off
echo ExamOffice setup - installing the one dependency (python-docx)...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
echo Done. Double-click run.bat to start ExamOffice.
pause
