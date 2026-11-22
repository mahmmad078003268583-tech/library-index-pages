@echo off
rem ============================================================
rem  Auto-start the library server at Windows logon.
rem  Run this file ONCE on every computer where the server
rem  should start by itself (copy the folder first, then run it).
rem ============================================================
setlocal
set "APPDIR=%~dp0"

if not exist "%APPDIR%autostart.ps1" goto :missing

powershell -NoProfile -ExecutionPolicy Bypass -File "%APPDIR%autostart.ps1"
echo.
pause
exit /b 0

:missing
echo.
echo   autostart.ps1 was not found in:
echo   %APPDIR%
echo.
echo   Keep this file together with autostart.ps1.
echo.
pause
exit /b 1
