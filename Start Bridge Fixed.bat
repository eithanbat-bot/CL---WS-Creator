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
echo Checking GitHub for the newest Creator bridge...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri '%UPDATER_URL%' -OutFile '%TEMP_UPDATER%' -UseBasicParsing -Headers @{'User-Agent'='CL-WS-Creator-Updater'} -TimeoutSec 60; exit 0 } catch { Write-Host '[CL-WS] Update download failed:'; Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 goto LOCAL_FALLBACK

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TEMP_UPDATER%" -Root "%ROOT%"
goto END

:LOCAL_FALLBACK
echo.
echo [CL-WS] Using the local updater as a fallback.
if exist "%ROOT%bridge\update-and-start.ps1" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%bridge\update-and-start.ps1" -Root "%ROOT%"
  goto END
)

echo [CL-WS] No local updater found. Starting the existing bridge.
if exist "%ROOT%bridge\server.ps1" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ROOT%bridge\server.ps1"

:END
echo.
echo Bridge stopped. Press any key to close.
pause >nul
