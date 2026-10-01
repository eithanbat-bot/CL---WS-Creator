<#
  CL - WS Creator - bridge updater / launcher
  - Downloads the newest bridge scripts from GitHub (main branch) on every start.
  - Validates every downloaded script (PowerShell syntax) BEFORE touching the installed copy.
  - Only replaces files that actually changed; keeps the previous version in bridge\backup.
  - Never touches local data: config.json, dxf-index.tsv, dxf-index-status.json.
  - Works offline: if GitHub is unreachable, starts the copy already installed.
  - If a freshly updated bridge crashes at startup, rolls back to the backup and starts that.
  To add a new bridge file in future, add its name to $Files below and push - nothing else.
#>
param([string]$Root)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$Repo = 'eithanbat-bot/CL---WS-Creator'
$Branch = 'main'
$Base = "https://raw.githubusercontent.com/$Repo/$Branch/bridge/"
$Files = @('server.ps1', 'create-sigmanest-ws.ps1', 'dxf-indexer.ps1', 'update-and-start.ps1')
$Port = 17832

function Say($m, $c = 'Gray') { Write-Host "[CL-WS] $m" -ForegroundColor $c }

# ---- resolve folders (strip stray quotes / trailing slash from the launcher) ----
if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = ($Root -replace '"', '').Trim().TrimEnd('\')
$BridgeDir = Join-Path $Root 'bridge'
$BackupDir = Join-Path $BridgeDir 'backup'
$Server = Join-Path $BridgeDir 'server.ps1'
New-Item -ItemType Directory -Force -Path $BridgeDir | Out-Null

function Get-Version($path) {
  try {
    if (-not (Test-Path -LiteralPath $path)) { return 'none' }
    $m = Select-String -LiteralPath $path -Pattern '\$BRIDGE_VERSION\s*=\s*''([^'']+)''' | Select-Object -First 1
    if ($m) { return $m.Matches[0].Groups[1].Value }
  } catch {}
  return 'unknown'
}
function Test-Syntax($path) {
  $tokens = $null; $errs = $null
  [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errs) | Out-Null
  return (-not $errs -or $errs.Count -eq 0)
}
function Get-Hash($path) {
  if (Test-Path -LiteralPath $path) { return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
  return ''
}

Say "Installed bridge version: $(Get-Version $Server)"

# ---- 1. download + validate everything into a temp folder first ----
$updated = $false
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('clws-update-' + [Guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Force -Path $tmp | Out-Null
  $stamp = [DateTime]::UtcNow.Ticks
  foreach ($f in $Files) {
    $dest = Join-Path $tmp $f
    Invoke-WebRequest -Uri ($Base + $f + '?t=' + $stamp) -OutFile $dest -UseBasicParsing -TimeoutSec 60 `
      -Headers @{ 'User-Agent' = 'CL-WS-Creator-Updater'; 'Cache-Control' = 'no-cache' }
    if (-not (Test-Path -LiteralPath $dest) -or (Get-Item -LiteralPath $dest).Length -lt 50) { throw "Downloaded $f is empty." }
    if ($f -like '*.ps1' -and -not (Test-Syntax $dest)) { throw "Downloaded $f failed the PowerShell syntax check." }
  }
  Say "Latest bridge version on GitHub: $(Get-Version (Join-Path $tmp 'server.ps1'))"

  # ---- 2. swap in only the files that changed (backup the old ones) ----
  $changed = @($Files | Where-Object { (Get-Hash (Join-Path $tmp $_)) -ne (Get-Hash (Join-Path $BridgeDir $_)) })
  if ($changed.Count -gt 0) {
    if (Test-Path -LiteralPath $BackupDir) { Remove-Item -LiteralPath $BackupDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
    foreach ($f in $changed) {
      $cur = Join-Path $BridgeDir $f
      if (Test-Path -LiteralPath $cur) { Copy-Item -LiteralPath $cur -Destination (Join-Path $BackupDir $f) -Force }
      Copy-Item -LiteralPath (Join-Path $tmp $f) -Destination $cur -Force
    }
    $updated = $true
    Say ("Updated: " + ($changed -join ', ')) 'Green'
  } else {
    Say 'Bridge is already up to date.' 'Green'
  }
} catch {
  Say "Could not update ($($_.Exception.Message)). Using the installed copy." 'Yellow'
} finally {
  if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

if (-not (Test-Path -LiteralPath $Server)) {
  Say 'No bridge is installed and GitHub could not be reached. Connect to the internet and run this again.' 'Red'
  exit 1
}

# ---- 3. stop any older bridge still holding the port, so the new code really runs ----
try {
  $self = $PID
  Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.ProcessId -ne $self -and $_.CommandLine -match 'bridge[\\/]server\.ps1' } |
    ForEach-Object { Say "Stopping older bridge (PID $($_.ProcessId))..." 'Yellow'; Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
  Start-Sleep -Milliseconds 400
} catch {}

# ---- 4. start the bridge; if a just-updated bridge dies immediately, roll back ----
function Start-Bridge {
  $sw = [Diagnostics.Stopwatch]::StartNew()
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Server | Out-Host
  $code = $LASTEXITCODE
  $sw.Stop()
  return @{ Code = $code; Seconds = $sw.Elapsed.TotalSeconds }
}

Say "Starting bridge $(Get-Version $Server) on http://127.0.0.1:$Port  (leave this window open while using Excel)" 'Cyan'
$r = Start-Bridge
if ($updated -and $r.Seconds -lt 20 -and (Test-Path -LiteralPath $BackupDir)) {
  Say 'The updated bridge stopped right after starting. Rolling back to the previous version...' 'Red'
  foreach ($f in $Files) {
    $b = Join-Path $BackupDir $f
    if (Test-Path -LiteralPath $b) { Copy-Item -LiteralPath $b -Destination (Join-Path $BridgeDir $f) -Force }
  }
  Say "Restarted previous version $(Get-Version $Server)." 'Yellow'
  Start-Bridge | Out-Null
}
