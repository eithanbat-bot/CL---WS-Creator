@echo off
cd /d "%~dp0"
echo.
echo Starting CL - WS Creator local bridge...
echo This version uses Windows PowerShell. Node.js is NOT required.
echo.

REM Always refresh the bridge scripts into the actual bridge folder.
REM This prevents an old copy in the project root from being executed.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $base='https://raw.githubusercontent.com/eithanbat-bot/CL---WS-Creator/main/'; New-Item -ItemType Directory -Force -Path '%~dp0bridge' | Out-Null; Invoke-WebRequest -UseBasicParsing -Uri ($base+'bridge/server.ps1') -OutFile '%~dp0bridgeserver.ps1'; Invoke-WebRequest -UseBasicParsing -Uri ($base+'bridge/create-sigmanest-ws.ps1') -OutFile '%~dp0bridgecreate-sigmanest-ws.ps1'; Write-Host 'Bridge scripts refreshed from GitHub.'"

if errorlevel 1 (
  echo.
  echo ERROR: Could not refresh bridge scripts from GitHub.
  echo The bridge was NOT started.
  echo.
  pause
  exit /b 1
)

REM Parse the bridge script before starting it. This catches syntax regressions
REM locally instead of starting a broken server.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$path='%~dp0bridgeserver.ps1'; $tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors) | Out-Null; if($errors.Count -gt 0){ Write-Host 'ERROR: server.ps1 failed PowerShell syntax validation.' -ForegroundColor Red; $errors | ForEach-Object { Write-Host $_.Message -ForegroundColor Red }; exit 1 }; Write-Host 'server.ps1 syntax check: OK.' -ForegroundColor Green"

if errorlevel 1 (
  echo.
  echo Bridge was NOT started because server.ps1 failed syntax validation.
  echo.
  pause
  exit /b 1
)

echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0bridgeserver.ps1"
echo.
echo Bridge stopped. Press any key to close this window.
pause >nul
