@echo off
cd /d "%~dp0"
echo.
echo Starting CL - WS Creator local bridge...
echo This version uses Windows PowerShell. Node.js is NOT required.
echo.

REM Refresh the two PowerShell bridge scripts from the public GitHub repository.
REM This prevents an older local copy from being used after the bridge is updated.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $base='https://raw.githubusercontent.com/eithanbat-bot/CL---WS-Creator/main/'; Invoke-WebRequest -UseBasicParsing -Uri ($base+'bridge/server.ps1') -OutFile '%~dp0bridgeserver.ps1'; Invoke-WebRequest -UseBasicParsing -Uri ($base+'bridge/create-sigmanest-ws.ps1') -OutFile '%~dp0bridgecreate-sigmanest-ws.ps1'; Write-Host 'Bridge scripts refreshed from GitHub.'"

if errorlevel 1 (
  echo.
  echo WARNING: Could not refresh bridge scripts from GitHub.
  echo The existing local scripts will be used.
  echo.
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0bridgeserver.ps1"
echo.
echo Bridge stopped. Press any key to close this window.
pause >nul
