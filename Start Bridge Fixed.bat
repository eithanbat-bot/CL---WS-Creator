@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ROOT=%~dp0"
set "TMP=%TEMP%\CL-WS-Creator-bootstrap"
set "RAW=https://raw.githubusercontent.com/eithanbat-bot/CL---WS-Creator/main/bridge/update-and-start.ps1"

echo ==============================================
echo CL - WS Creator - Bridge
echo ==============================================
echo.
echo This launcher downloads the current bridge updater from GitHub.
echo Local config and DXF index data are preserved.
echo.

if exist "%TMP%" del /f /q "%TMP%" >nul 2>&1
mkdir "%TMP%" >nul 2>&1

echo [CL-WS] Downloading current self-updater...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $urls=@('%RAW%','https://github.com/eithanbat-bot/CL---WS-Creator/raw/refs/heads/main/bridge/update-and-start.ps1'); $ok=$false; $errors=@(); foreach($u in $urls){try{Invoke-WebRequest -UseBasicParsing -TimeoutSec 120 -Uri $u -OutFile '%TMP%\update-and-start.ps1' -Headers @{'User-Agent'='CL-WS-Creator-Bootstrap';'Cache-Control'='no-cache'}; if((Test-Path -LiteralPath '%TMP%\update-and-start.ps1') -and (Get-Item -LiteralPath '%TMP%\update-and-start.ps1').Length -gt 100){$ok=$true;break}}catch{$errors+=($u+': '+$_.Exception.Message)}}; if(-not $ok){Write-Warning ('Could not download the current updater: '+($errors -join ' | '))} else {try{$text=Get-Content -LiteralPath '%TMP%\update-and-start.ps1' -Raw -Encoding UTF8; [scriptblock]::Create($text)|Out-Null} catch {Remove-Item '%TMP%\update-and-start.ps1' -Force -ErrorAction SilentlyContinue; Write-Warning ('Downloaded bridge updater failed PowerShell syntax validation: '+$_.Exception.Message)}}"
 
set "UPDATER=%TMP%\update-and-start.ps1"
if not exist "%UPDATER%" set "UPDATER=%ROOT%bridge\update-and-start.ps1"
if not exist "%UPDATER%" (
  echo [CL-WS] No usable bridge updater is available locally or from GitHub.
  goto FAIL
)

echo [CL-WS] Starting self-updater...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%UPDATER%" -Root "%ROOT%"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" goto FAIL
if exist "%TMP%" rmdir /s /q "%TMP%" >nul 2>&1
exit /b 0

:FAIL
echo.
echo [CL-WS] Bridge update/start FAILED.
echo [CL-WS] Check the messages above and run this launcher again after correcting the reported issue.
echo.
if exist "%TMP%" rmdir /s /q "%TMP%" >nul 2>&1
pause
exit /b 1
