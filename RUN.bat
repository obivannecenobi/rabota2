@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
pushd "%~dp0"

set "VENV=.venv"
set "REQS=requirements.txt"
set "APP=app.main"
set "PY=%VENV%\Scripts\python.exe"
set "PYW=%VENV%\Scripts\pythonw.exe"

if not exist "%PY%" if not exist "%PYW%" (
  where py >nul 2>nul && (py -3 -m venv "%VENV%") || (python -m venv "%VENV%")
)

if exist "%PY%" (
  "%PY%" -m pip install --upgrade pip
  if errorlevel 1 goto err
  "%PY%" -m pip install -r "%REQS%"
  if errorlevel 1 goto err
  if exist "%PYW%" (
    start "" "%PYW%" -m %APP%
  ) else (
    start "" "%PY%" -m %APP%
  )
) else if exist "%PYW%" (
  "%PYW%" -m pip install --upgrade pip
  if errorlevel 1 goto err
  "%PYW%" -m pip install -r "%REQS%"
  if errorlevel 1 goto err
  start "" "%PYW%" -m %APP%
) else (
  echo [ERR] Не найден Python. Установи Python 3.10+ и попробуй снова.
  pause
  goto :eof
)

popd
exit /b 0

:err
echo [ERR] Ошибка при установке зависимостей.
pause
exit /b 1
