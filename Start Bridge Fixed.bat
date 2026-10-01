@echo off
setlocal
cd /d "%~dp0"

echo ==============================================
echo CL - WS Creator local bridge
echo ==============================================
echo Starting PowerShell bridge from:
echo %~dp0bridge\server.ps1
echo.
echo Leave this window open while using the Excel add-in.
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0bridge\server.ps1"

echo.
echo Bridge stopped. Press any key to close.
pause >nul
