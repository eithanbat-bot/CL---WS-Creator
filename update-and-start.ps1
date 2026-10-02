<#
  CL - WS Creator - self-updating bridge launcher.
  Run the same Start Bridge Fixed.bat every time.
  The launcher downloads the current bridge runtime from GitHub, validates
  PowerShell files before installation, preserves generated local DXF data,
  updates a trusted-folder manifest when one exists, and then starts the bridge.
#>
param([string]$Root)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$Repo = 'eithanbat-bot/CL---WS-Creator'
$Branch = 'main'
$ApiTree = "https://api.github.com/repos/$Repo/git/trees/$Branch?recursive=1"
$RawBase = "https://raw.githubusercontent.com/$Repo/$Branch/"
$Port = 17832

function Say([string]$Message,[string]$Color='Gray') {
  Write-Host "[CL-WS] $Message" -ForegroundColor $Color
}

function Get-Version([string]$Path) {
  try {
    if(-not(Test-Path -LiteralPath $Path)){return 'none'}
    $m=Select-String -LiteralPath $Path -Pattern '\$BRIDGE_VERSION\s*=\s*''([^'']+)''' | Select-Object -First 1
    if($m){return $m.Matches[0].Groups[1].Value}
  }catch{}
  return 'unknown'
}

function Test-Syntax([string]$Path) {
  $tokens=$null
  $errors=$null
  [System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)|Out-Null
  return ($null -eq $errors -or $errors.Count -eq 0)
}

function Get-Hash([string]$Path) {
  if(Test-Path -LiteralPath $Path){
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
  }
  return ''
}

function Download-Json([string]$Url) {
  Invoke-RestMethod -Uri $Url -UseBasicParsing -TimeoutSec 60 -Headers @{
    'User-Agent'='CL-WS-Creator-Updater'
    'Accept'='application/vnd.github+json'
    'Cache-Control'='no-cache'
  }
}

function Download-File([string]$Url,[string]$Destination) {
  $parent=Split-Path -Parent $Destination
  if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
  $temp=$Destination+'.download'
  try{
    Invoke-WebRequest -Uri $Url -OutFile $temp -UseBasicParsing -TimeoutSec 60 -Headers @{
      'User-Agent'='CL-WS-Creator-Updater'
      'Cache-Control'='no-cache'
    }
    if(-not(Test-Path -LiteralPath $temp)){throw 'Download did not produce a file.'}
    if((Get-Item -LiteralPath $temp).Length -lt 20){throw 'Downloaded file is empty.'}
    Move-Item -LiteralPath $temp -Destination $Destination -Force
  }finally{
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
  }
}

function Stop-OlderBridges {
  param([string]$ServerPath)
  try{
    $pattern=[regex]::Escape($ServerPath)
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
      Where-Object {
        $_.ProcessId -ne $PID -and
        ([string]$_.CommandLine -match $pattern)
      } |
      ForEach-Object {
        Say "Stopping older Creator bridge (PID $($_.ProcessId))..." 'Yellow'
        Stop-Process -Id ([int]$_.ProcessId) -Force -ErrorAction SilentlyContinue
      }
  }catch{}
  Start-Sleep -Milliseconds 400
}

if([string]::IsNullOrWhiteSpace($Root)){
  $Root=Split-Path -Parent $PSScriptRoot
}
$Root=$Root.Trim().Trim('"').Trim("'")
try{
  $Root=(Resolve-Path -LiteralPath $Root -ErrorAction Stop).Path
}catch{
  throw "Invalid Creator installation folder: $Root"
}

$BridgeDir=Join-Path $Root 'bridge'
$BackupDir=Join-Path $BridgeDir 'backup'
$Server=Join-Path $BridgeDir 'server.ps1'
New-Item -ItemType Directory -Path $BridgeDir -Force|Out-Null

Say "Installed bridge version: $(Get-Version $Server)"
$tmpRoot=Join-Path ([IO.Path]::GetTempPath()) ('clwsc-update-'+[Guid]::NewGuid().ToString('N'))

try{
  New-Item -ItemType Directory -Path $tmpRoot -Force|Out-Null
  Say 'Checking GitHub for the newest bridge files...'

  $tree=Download-Json $ApiTree
  $files=@(
    $tree.tree |
      Where-Object {
        $_.type -eq 'blob' -and
        $_.path -like 'bridge/*' -and
        $_.path -notmatch '^bridge/(config\.json|dxf-index\.json|dxf-index-status\.json)$' -and
        $_.path -notmatch '^bridge/dxf-index/' -and
        $_.path -notmatch '^bridge/backup/'
      } |
      ForEach-Object { [string]$_.path.Substring(7) }
  )

  if($files.Count -eq 0){throw 'GitHub returned no bridge runtime files.'}

  foreach($relative in $files){
    $downloadPath=Join-Path $tmpRoot $relative
    $remoteUrl=$RawBase+'bridge/'+$relative.Replace('\','/')
    Say "Checking bridge\$relative..."
    Download-File $remoteUrl $downloadPath

    if([IO.Path]::GetExtension($relative).ToLowerInvariant() -eq '.ps1'){
      if(-not(Test-Syntax $downloadPath)){
        throw "Downloaded $relative failed the PowerShell syntax check."
      }
    }
  }

  $remoteServer=Join-Path $tmpRoot 'server.ps1'
  if(-not(Test-Path -LiteralPath $remoteServer)){throw 'GitHub bridge does not contain server.ps1.'}
  Say "Latest bridge version on GitHub: $(Get-Version $remoteServer)" 'Cyan'

  $changed=@(
    $files | Where-Object {
      $remote=Join-Path $tmpRoot $_
      $local=Join-Path $BridgeDir $_
      (Get-Hash $remote) -ne (Get-Hash $local)
    }
  )

  if($changed.Count -gt 0){
    if(Test-Path -LiteralPath $BackupDir){Remove-Item -LiteralPath $BackupDir -Recurse -Force}
    New-Item -ItemType Directory -Path $BackupDir -Force|Out-Null

    foreach($relative in $changed){
      $local=Join-Path $BridgeDir $relative
      if(Test-Path -LiteralPath $local){
        $backup=Join-Path $BackupDir $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $backup) -Force|Out-Null
        Copy-Item -LiteralPath $local -Destination $backup -Force
      }
    }

    try{
      foreach($relative in $changed){
        $local=Join-Path $BridgeDir $relative
        New-Item -ItemType Directory -Path (Split-Path -Parent $local) -Force|Out-Null
        Copy-Item -LiteralPath (Join-Path $tmpRoot $relative) -Destination $local -Force
      }
      Say ("Updated bridge files: "+($changed -join ', ')) 'Green'
    }catch{
      Say 'Bridge replacement failed; restoring the previous bridge files...' 'Red'
      foreach($relative in $changed){
        $backup=Join-Path $BackupDir $relative
        $local=Join-Path $BridgeDir $relative
        if(Test-Path -LiteralPath $backup){
          Copy-Item -LiteralPath $backup -Destination $local -Force
        }
      }
      throw
    }
  }else{
    Say 'Bridge files are already current.' 'Green'
  }

  $localManifest=Join-Path $Root 'manifest.xml'
  if(Test-Path -LiteralPath $localManifest){
    $remoteManifest=Join-Path $tmpRoot 'manifest.xml'
    Say 'Checking local manifest.xml...'
    Download-File ($RawBase+'manifest.xml') $remoteManifest
    try{
      [xml](Get-Content -Raw -LiteralPath $remoteManifest)|Out-Null
    }catch{
      throw 'Downloaded manifest.xml failed XML validation.'
    }

    if((Get-Hash $remoteManifest) -ne (Get-Hash $localManifest)){
      Copy-Item -LiteralPath $remoteManifest -Destination $localManifest -Force
      Say 'Updated local manifest.xml.' 'Green'
    }else{
      Say 'Local manifest.xml is already current.' 'Green'
    }
  }
}catch{
  Say "Could not update ($($_.Exception.Message)). Using the installed copy." 'Yellow'
}finally{
  Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if(-not(Test-Path -LiteralPath $Server)){
  Say 'No working Creator bridge is installed.' 'Red'
  exit 1
}

Stop-OlderBridges -ServerPath $Server
Say "Starting bridge $(Get-Version $Server) on http://127.0.0.1:$Port (leave this window open while using Excel)" 'Cyan'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Server
