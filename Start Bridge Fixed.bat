@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ROOT=%~dp0"
set "BRIDGE=%ROOT%bridge"
set "TMP=%TEMP%\CL-WS-Creator-runtime"
set "RAW=https://raw.githubusercontent.com/eithanbat-bot/CL---WS-Creator/main/bridge"

echo ==============================================
echo CL - WS Creator - Bridge
echo ==============================================
echo.
echo This launcher installs the current bridge from GitHub.
echo Local config and DXF index data are preserved.
echo.

if not exist "%BRIDGE%" mkdir "%BRIDGE%" >nul 2>&1
if exist "%TMP%" rmdir /s /q "%TMP%" >nul 2>&1
mkdir "%TMP%" >nul 2>&1

echo [CL-WS] Stopping any existing Creator bridge using this install...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $target=[IO.Path]::GetFullPath('%BRIDGE%\server.ps1'); Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" | Where-Object { $_.ProcessId -ne $PID -and [regex]::IsMatch([string]$_.CommandLine,[regex]::Escape($target)) } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }; Start-Sleep -Milliseconds 500"
echo.
echo [CL-WS] Downloading server.ps1...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Invoke-WebRequest -UseBasicParsing -TimeoutSec 90 -Uri '%RAW%/server.ps1' -OutFile '%TMP%\server.ps1'; $t=Get-Content -Raw '%TMP%\server.ps1'; if($t -notmatch '\$BRIDGE_VERSION\s*=\s*''2\.3\.0'''){throw 'Downloaded server.ps1 is not bridge version 2.3.0.'}; $x=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('%TMP%\server.ps1',[ref]$x,[ref]$e)|Out-Null;if($e.Count -gt 0){throw 'Downloaded server.ps1 failed PowerShell syntax validation.'}"
if errorlevel 1 goto FAIL

echo [CL-WS] Downloading dxf-indexer.ps1...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Invoke-WebRequest -UseBasicParsing -TimeoutSec 90 -Uri '%RAW%/dxf-indexer.ps1' -OutFile '%TMP%\dxf-indexer.ps1'; $x=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('%TMP%\dxf-indexer.ps1',[ref]$x,[ref]$e)|Out-Null;if($e.Count -gt 0){throw 'Downloaded dxf-indexer.ps1 failed PowerShell syntax validation.'}"
if errorlevel 1 goto FAIL

echo [CL-WS] Downloading create-sigmanest-ws.ps1...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Invoke-WebRequest -UseBasicParsing -TimeoutSec 90 -Uri '%RAW%/create-sigmanest-ws.ps1' -OutFile '%TMP%\create-sigmanest-ws.ps1'; $x=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('%TMP%\create-sigmanest-ws.ps1',[ref]$x,[ref]$e)|Out-Null;if($e.Count -gt 0){throw 'Downloaded create-sigmanest-ws.ps1 failed PowerShell syntax validation.'}"
if errorlevel 1 goto FAIL

echo [CL-WS] Installing current bridge...
copy /y "%TMP%\server.ps1" "%BRIDGE%\server.ps1" >nul
copy /y "%TMP%\dxf-indexer.ps1" "%BRIDGE%\dxf-indexer.ps1" >nul
copy /y "%TMP%\create-sigmanest-ws.ps1" "%BRIDGE%\create-sigmanest-ws.ps1" >nul

echo [CL-WS] Confirming installed version...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$t=Get-Content -Raw '%BRIDGE%\server.ps1'; if($t -notmatch '\$BRIDGE_VERSION\s*=\s*''2\.3\.0'''){throw 'Installed bridge is not version 2.3.0.'}; Write-Host '[CL-WS] Installed bridge version: 2.3.0' -ForegroundColor Green"
if errorlevel 1 goto FAIL

echo.
echo [CL-WS] Starting bridge 2.3.0 on http://127.0.0.1:17832
echo [CL-WS] Leave this window open while using Excel.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BRIDGE%\server.ps1"
goto END

:FAIL
echo.
echo [CL-WS] Bridge update FAILED. No new runtime files were started.
echo [CL-WS] Your existing local bridge files were not deliberately deleted.
echo.
if exist "%TMP%" rmdir /s /q "%TMP%" >nul 2>&1
pause
exit /b 1

:END
if exist "%TMP%" rmdir /s /q "%TMP%" >nul 2>&1
echo.
echo Bridge stopped. Press any key to close.
pause >nul
