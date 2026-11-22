@echo off
title Library Index Server
cd /d "%~dp0"
echo.
echo   Starting Library Index Server...
echo   Close this window to stop the server.
echo.

rem Find Python: full paths first, then PATH, then py launcher
set "PY="
if exist "C:\Program Files\Python313\python.exe" set "PY=C:\Program Files\Python313\python.exe"
if not defined PY if exist "C:\Program Files\Python312\python.exe" set "PY=C:\Program Files\Python312\python.exe"
if not defined PY if exist "C:\Program Files\Python311\python.exe" set "PY=C:\Program Files\Python311\python.exe"
if defined PY goto run
where python >nul 2>nul
if not errorlevel 1 set "PY=python"
if defined PY goto run
set "PY=py"

:run
echo   Using: %PY%
rem Run: "%PY%" server.py   (extra arguments are passed through, e.g. --no-browser)
"%PY%" server.py %*
if errorlevel 1 (
  echo.
  echo   Failed to start. Install Python and required packages:
  echo      pip install openpyxl qrcode pillow
  echo.
  pause
)
