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
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $api=Invoke-RestMethod -UseBasicParsing -TimeoutSec 120 -Uri 'https://api.github.com/repos/eithanbat-bot/CL---WS-Creator/contents/bridge/update-and-start.ps1?ref=main' -Headers @{'User-Agent'='CL-WS-Creator-Bootstrap';'Cache-Control'='no-cache'}; Invoke-WebRequest -UseBasicParsing -TimeoutSec 120 -Uri '%RAW%' -OutFile '%TMP%\update-and-start.ps1'; $b=[IO.File]::ReadAllBytes('%TMP%\update-and-start.ps1'); $h=[Text.Encoding]::ASCII.GetBytes(('blob '+$b.Length+[char]0)); $all=New-Object byte[] ($h.Length+$b.Length); [Array]::Copy($h,0,$all,0,$h.Length); [Array]::Copy($b,0,$all,$h.Length,$b.Length); $s=[Security.Cryptography.SHA1]::Create(); try{$actual=(($s.ComputeHash($all)|ForEach-Object{$_.ToString('x2')})-join '')}finally{$s.Dispose()}; if($actual.ToLowerInvariant() -ne ([string]$api.sha).ToLowerInvariant()){throw 'Downloaded bridge updater failed Git blob integrity validation.'}; $x=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('%TMP%\update-and-start.ps1',[ref]$x,[ref]$e)|Out-Null;if($e.Count -gt 0){throw 'Downloaded bridge updater failed PowerShell syntax validation.'}"
if errorlevel 1 goto FAIL

echo [CL-WS] Starting verified self-updater...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TMP%\update-and-start.ps1" -Root "%ROOT%"
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
