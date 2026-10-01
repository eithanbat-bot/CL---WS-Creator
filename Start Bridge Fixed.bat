@echo off
setlocal
cd /d "%~dp0"

set "ROOT=%~dp0"
set "UPDATER_URL=https://raw.githubusercontent.com/eithanbat-bot/CL---WS-Creator/main/bridge/update-and-start.ps1"
set "TEMP_UPDATER=%TEMP%\CL-WS-Creator-update-and-start.ps1"

echo ==============================================
echo CL - WS Creator - Self Updating Bridge
echo ==============================================
echo.
echo This launcher downloads the newest Creator bridge
echo files automatically before starting the bridge.
echo Local DXF index/config data is preserved.
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; try { Invoke-WebRequest -Uri '%UPDATER_URL%' -OutFile '%TEMP_UPDATER%' -UseBasicParsing -Headers @{'User-Agent'='CL-WS-Creator-Updater'} -TimeoutSec 60; powershell.exe -NoProfile -ExecutionPolicy Bypass -File '%TEMP_UPDATER%' -Root '%ROOT%' } catch { Write-Host '[CL-WS] Could not download the latest updater.'; Write-Host $_.Exception.Message; if (Test-Path '%ROOT%bridge\update-and-start.ps1') { powershell.exe -NoProfile -ExecutionPolicy Bypass -File '%ROOT%bridge\update-and-start.ps1' -Root '%ROOT%' } else { powershell.exe -NoProfile -ExecutionPolicy Bypass -File '%ROOT%bridge\server.ps1' } }"

echo.
echo Bridge stopped. Press any key to close.
pause >nul
