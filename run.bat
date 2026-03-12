@echo off
echo ==========================================
echo   Requirements Tracker
echo ==========================================
echo.

REM Check if dependencies are installed
pip show PyMuPDF >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r "%~dp0requirements.txt"
    echo.
)

python "%~dp0requirements_tracker.py"
pause