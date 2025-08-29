@echo off
cd /d %~dp0

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not added to PATH.
    exit /b 1
)

if not exist ".\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Please create it with: python -m venv .venv
    exit /b 1
)

".\.venv\Scripts\python" run.py
