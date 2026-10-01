<#
  CL - WS Creator - self-updating bridge launcher
  Downloads and validates every bridge file currently present under bridge/ on GitHub.
  Machine-generated files are preserved:
    config.json
    dxf-index.json
    dxf-index-status.json
    dxf-index/
  Local manifest is refreshed when it exists.
  The user only needs to run the same BAT again for future bridge updates.
#>
param([string]$Root)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$Repo = 'eithanbat-bot/CL---WS-Creator'
$Branch = 'main'
$ApiTree = "https://api.github.com/repos/$Repo/git/trees/$Branch?recursive=1"
$RawBase = "https://raw.githubusercontent.com/$Repo/$Branch/"
$Port = 17832

function Say($m, $c = 'Gray') { Write-Host "[CL-WS] $m" -ForegroundColor $c }

if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = ($Root -replace '"', '').Trim()
try { $Root = (Resolve-Path -LiteralPath $Root -ErrorAction Stop).Path } catch {
  throw "Invalid Creator installation folder: $Root"
}

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
  $tokens = $null
  $errs = $null
  [System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errs) | Out-Null
  return (-not $errs -or $errs.Count -eq 0)
}

function Get-Hash($path) {
  if (Test-Path -LiteralPath $path) { return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
  return ''
}

function Download-Text($url) {
  Invoke-RestMethod -Uri $url -UseBasicParsing -Headers @{
    'User-Agent'='CL-WS-Creator-Updater'
    'Accept'='application/vnd.github+json'
  } -TimeoutSec 60
}

function Download-File($url,$dest) {
  $parent=Split-Path -Parent $dest
  if($parent){New-Item -ItemType Directory -Force -Path $parent|Out-Null}
  $tmp=$dest+'.update'
  try{
    Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing -TimeoutSec 60 -Headers @{
      'User-Agent'='CL-WS-Creator-Updater'
      'Cache-Control'='no-cache'
    }
    if(-not(Test-Path -LiteralPath $tmp) -or (Get-Item -LiteralPath $tmp).Length -lt 20){
      throw 'Downloaded file is empty.'
    }
    Move-Item -LiteralPath $tmp -Destination $dest -Force
  }finally{
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  }
}

function Stop-OlderBridges(){
  try{
    $serverPattern=[regex]::Escape($Server)
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
      Where-Object {
        $_.ProcessId -ne $PID -and
        ([string]$_.CommandLine -match $serverPattern)
      } |
      ForEach-Object {
        Say "Stopping older Creator bridge (PID $($_.ProcessId))..." 'Yellow'
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
      }
  }catch{}
  Start-Sleep -Milliseconds 400
}

Say "Installed bridge version: $(Get-Version $Server)"

$updated=$false
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('clws-update-'+[Guid]::NewGuid().ToString('N'))

try{
  New-Item -ItemType Directory -Force -Path $tmp|Out-Null
  Say 'Checking GitHub for the newest bridge files...'

  $tree=Download-Text $ApiTree
  $files=@($tree.tree|Where-Object {
    $_.type -eq 'blob' -and
    $_.path -like 'bridge/*' -and
    $_.path -notmatch '^bridge/(config\.json|dxf-index\.json|dxf-index-status\.json)$' -and
    $_.path -notmatch '^bridge/dxf-index/' -and
    $_.path -notmatch '^bridge/backup/'
  }|ForEach-Object {$_.path.Substring(7)})

  if(-not $files.Count){throw 'GitHub returned no bridge runtime files.'}

  foreach($rel in $files){
    $dest=Join-Path $tmp $rel
    $url=$RawBase+'bridge/'+$rel.Replace('\\','/')
    Say "Downloading bridge\$rel..."
    Download-File $url $dest
    if([IO.Path]::GetExtension($rel).ToLowerInvariant() -eq '.ps1' -and -not(Test-Syntax $dest)){
      throw "Downloaded $rel failed the PowerShell syntax check."
    }
  }

  Say "Latest bridge version on GitHub: $(Get-Version (Join-Path $tmp 'server.ps1'))"

  $changed=@($files|Where-Object {
    $remote=Join-Path $tmp $_
    $local=Join-Path $BridgeDir $_
    (Get-Hash $remote) -ne (Get-Hash $local)
  })

  if($changed.Count -gt 0){
    if(Test-Path -LiteralPath $BackupDir){Remove-Item -LiteralPath $BackupDir -Recurse -Force}
    New-Item -ItemType Directory -Force -Path $BackupDir|Out-Null

    foreach($rel in $changed){
      $local=Join-Path $BridgeDir $rel
      $backup=Join-Path $BackupDir $rel
      if(Test-Path -LiteralPath $local){
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $backup)|Out-Null
        Copy-Item -LiteralPath $local -Destination $backup -Force
      }
    }

    try{
      foreach($rel in $changed){
        $local=Join-Path $BridgeDir $rel
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $local)|Out-Null
        Copy-Item -LiteralPath (Join-Path $tmp $rel) -Destination $local -Force
      }
      $updated=$true
      Say ("Updated bridge files: "+($changed -join ', ')) 'Green'
    }catch{
      Say 'Bridge update failed during file replacement. Restoring previous files...' 'Red'
      foreach($rel in $changed){
        $backup=Join-Path $BackupDir $rel
        $local=Join-Path $BridgeDir $rel
        if(Test-Path -LiteralPath $backup){Copy-Item -LiteralPath $backup -Destination $local -Force}
      }
      throw
    }
  }else{
    Say 'Bridge is already up to date.' 'Green'
  }

  # Synchronize the trusted-folder manifest independently from bridge files.
  $localManifest=Join-Path $Root 'manifest.xml'
  if(Test-Path -LiteralPath $localManifest){
    $remoteManifest=Join-Path $tmp 'manifest.xml'
    Say 'Checking local manifest.xml...'
    Download-File ($RawBase+'manifest.xml') $remoteManifest
    try{ [xml](Get-Content -Raw -LiteralPath $remoteManifest)|Out-Null }
    catch{ throw 'Downloaded manifest.xml failed XML validation.' }

    if((Get-Hash $remoteManifest) -ne (Get-Hash $localManifest)){
      Copy-Item -LiteralPath $remoteManifest -Destination $localManifest -Force
      Say 'Updated local manifest.xml.' 'Green'
    }else{
      Say 'Local manifest.xml is already current.' 'Green'
    }
  }
  }else{
    Say 'Bridge is already up to date.' 'Green'
  }
}catch{
  Say "Could not update ($($_.Exception.Message)). Using the installed copy." 'Yellow'
}finally{
  Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if(-not(Test-Path -LiteralPath $Server)){
  Say 'No working bridge is installed. Run this launcher again while connected to GitHub.' 'Red'
  exit 1
}

Stop-OlderBridges
Say "Starting bridge $(Get-Version $Server) on http://127.0.0.1:$Port  (leave this window open while using Excel)" 'Cyan'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Server
