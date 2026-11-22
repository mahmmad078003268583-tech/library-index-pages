@echo off
rem ============================================================
rem  Remove auto-start: the server will no longer start by
rem  itself when Windows opens.  Run it on each computer where
rem  you want to cancel the automatic start.
rem ============================================================
setlocal
set "APPDIR=%~dp0"

if not exist "%APPDIR%autostart.ps1" goto :missing

powershell -NoProfile -ExecutionPolicy Bypass -File "%APPDIR%autostart.ps1" -Remove
echo.
pause
exit /b 0

:missing
echo.
echo   autostart.ps1 was not found in:
echo   %APPDIR%
echo.
pause
exit /b 1
