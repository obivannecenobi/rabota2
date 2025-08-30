@echo off
cd /d %~dp0

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not added to PATH.
    exit /b 1
)

rem Create virtual environment if it does not exist
if not exist ".\.venv\Scripts\python.exe" (
    echo Virtual environment not found. Creating...
    python -m venv .venv
)

if not exist ".\.venv\Scripts\python.exe" (
    echo Failed to create virtual environment or it is inaccessible.
    exit /b 1
)

".\.venv\Scripts\python" -m pip install --upgrade -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install dependencies.
    exit /b %errorlevel%
)

".\.venv\Scripts\python" app\main.py
if %errorlevel% neq 0 (
    echo Application exited with code %errorlevel%.
    exit /b %errorlevel%
)
