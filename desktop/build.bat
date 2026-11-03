@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo   ExamOffice - Windows Executable Build Script
echo ===================================================
echo.

cd /d "%~dp0"

:: 1. Ensure Python is available
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not found in PATH.
    echo Please install Python 3.10+ from python.org and tick "Add to PATH".
    pause
    exit /b 1
)

:: 2. Ensure virtual environment exists
if not exist "..\.venv\Scripts\python.exe" (
    echo [*] Creating virtual environment in ..\.venv ...
    python -m venv ..\.venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

set "VENV_PY=..\.venv\Scripts\python.exe"

:: 3. Install requirements and pyinstaller
echo [*] Installing requirements and PyInstaller...
"!VENV_PY!" -m pip install --upgrade pip
"!VENV_PY!" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: 4. Clean previous build artifacts
echo [*] Cleaning previous build artifacts...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

:: 5. Build with PyInstaller
echo [*] Building ExamOffice...
if "%~1"=="--onedir" (
    echo [*] Building onedir mode...
    "!VENV_PY!" -m PyInstaller --clean --noconfirm --windowed --onedir --name ExamOffice --icon icon.ico ^
        --collect-data docx --collect-data lxml ^
        --collect-submodules docx --collect-submodules lxml ^
        --hidden-import paper_loader --hidden-import docx_gen --hidden-import marker --hidden-import store ^
        --add-data "papers;papers" --add-data "icon.ico;." exam_office.py
) else (
    echo [*] Building onefile mode using ExamOffice.spec...
    "!VENV_PY!" -m PyInstaller --clean --noconfirm ExamOffice.spec
)

if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ===================================================
echo [SUCCESS] Build completed!
if exist "dist\ExamOffice.exe" (
    echo Binary: %~dp0dist\ExamOffice.exe
) else if exist "dist\ExamOffice\ExamOffice.exe" (
    echo Binary: %~dp0dist\ExamOffice\ExamOffice.exe
)
echo ===================================================
echo.
pause
