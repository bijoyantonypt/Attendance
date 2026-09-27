@echo off
title ESSL Attendance Report Generator
cd /d "%~dp0"

echo ============================================
echo   ESSL X990 - Local Attendance Report Tool
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from https://python.org and try again.
    pause
    exit /b
)

REM Create virtual environment if it doesn't exist
if not exist "venv\" (
    echo Creating virtual environment...
    python -m venv venv
)

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Installing/updating required packages...
pip install -r requirements.txt --quiet

echo.
echo Running report generator...
echo.
python generate_report.py

echo.
echo ============================================
echo Done. Check the "reports" folder for your Excel file.
echo ============================================
pause
