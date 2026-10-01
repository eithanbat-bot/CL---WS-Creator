@echo off
cd /d "%~dp0"
echo.
echo Starting CL - WS Creator local bridge...
echo This version uses Windows PowerShell. Node.js is NOT required.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0bridge\server.ps1"
echo.
echo Bridge stopped. Press any key to close this window.
pause >nul
