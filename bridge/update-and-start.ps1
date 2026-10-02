<#
  CL - WS Creator - self-updating bridge launcher.
  Start Bridge Fixed.bat downloads this launcher first, then this launcher
  fetches every bridge runtime file from the GitHub main branch, validates the
  Git blob SHA, validates PowerShell/XML syntax, installs atomically with rollback,
  preserves generated local DXF state/config, updates a local manifest when present,
  and starts the installed bridge.
#>
param([string]$Root)

$ErrorActionPreference='Stop'
try{[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12}catch{}

$Repo='eithanbat-bot/CL---WS-Creator'
$Branch='main'
$ApiTree="https://api.github.com/repos/$Repo/git/trees/$Branch?recursive=1"
$RawBase="https://raw.githubusercontent.com/$Repo/$Branch/"
$Port=17832

function Say([string]$Message,[string]$Color='Gray'){Write-Host "[CL-WS] $Message" -ForegroundColor $Color}

function Get-Version([string]$Path){
  try{
    if(-not(Test-Path -LiteralPath $Path)){return 'none'}
    $m=Select-String -LiteralPath $Path -Pattern '\$BRIDGE_VERSION\s*=\s*''([^'']+)'''|Select-Object -First 1
    if($m){return $m.Matches[0].Groups[1].Value}
  }catch{}
  return 'unknown'
}

function Test-Syntax([string]$Path){
  $tokens=$null;$errors=$null
  [System.Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)|Out-Null
  return ($null -eq $errors -or $errors.Count -eq 0)
}

function Test-Xml([string]$Path){
  try{[xml](Get-Content -Raw -LiteralPath $Path)|Out-Null;return $true}catch{return $false}
}

function Get-GitBlobSha1([string]$Path){
  $bytes=[IO.File]::ReadAllBytes($Path)
  $header=[Text.Encoding]::ASCII.GetBytes(('blob '+$bytes.Length+[char]0))
  $all=New-Object byte[] ($header.Length+$bytes.Length)
  [Array]::Copy($header,0,$all,0,$header.Length)
  [Array]::Copy($bytes,0,$all,$header.Length,$bytes.Length)
  $sha1=[Security.Cryptography.SHA1]::Create()
  try{return (($sha1.ComputeHash($all)|ForEach-Object{$_.ToString('x2')})-join '')}finally{$sha1.Dispose()}
}

function Get-Hash([string]$Path){
  if(Test-Path -LiteralPath $Path){return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash}
  return ''
}

function Invoke-Json([string]$Url){
  Invoke-RestMethod -Uri $Url -UseBasicParsing -TimeoutSec 60 -Headers @{
    'User-Agent'='CL-WS-Creator-Updater'
    'Accept'='application/vnd.github+json'
    'Cache-Control'='no-cache'
  }
}

function Download-Verified([string]$Url,[string]$Destination,[string]$ExpectedSha){
  $parent=Split-Path -Parent $Destination
  if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
  $tmp=$Destination+'.download'
  try{
    Invoke-WebRequest -Uri $Url -OutFile $tmp -UseBasicParsing -TimeoutSec 120 -Headers @{
      'User-Agent'='CL-WS-Creator-Updater'
      'Cache-Control'='no-cache'
    }
    if(-not(Test-Path -LiteralPath $tmp)){throw 'Download did not produce a file.'}
    if((Get-Item -LiteralPath $tmp).Length -lt 2){throw 'Downloaded file is empty.'}
    if($ExpectedSha){
      $actual=Get-GitBlobSha1 $tmp
      if($actual.ToLowerInvariant() -ne $ExpectedSha.ToLowerInvariant()){
        throw "Git blob SHA mismatch. Expected $ExpectedSha but downloaded $actual."
      }
    }
    Move-Item -LiteralPath $tmp -Destination $Destination -Force
  }finally{Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}
}

function Stop-CreatorProcesses([string]$ServerPath){
  try{
    $serverPattern=[regex]::Escape([IO.Path]::GetFullPath($ServerPath))
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue|
      Where-Object{
        $_.ProcessId -ne $PID -and (
          [string]$_.CommandLine -match $serverPattern -or
          [string]$_.CommandLine -match 'dxf-indexer\.ps1'
        )
      }|
      ForEach-Object{Say "Stopping existing Creator process PID $($_.ProcessId)...";Stop-Process -Id ([int]$_.ProcessId) -Force -ErrorAction SilentlyContinue}
  }catch{}
  Start-Sleep -Milliseconds 600
}

if([string]::IsNullOrWhiteSpace($Root)){$Root=Split-Path -Parent $PSScriptRoot}
$Root=$Root.Trim().Trim('"').Trim("'")
try{$Root=(Resolve-Path -LiteralPath $Root -ErrorAction Stop).Path}catch{throw "Invalid Creator installation folder: $Root"}

$BridgeDir=Join-Path $Root 'bridge'
$Server=Join-Path $BridgeDir 'server.ps1'
$BackupDir=Join-Path $BridgeDir 'backup'
New-Item -ItemType Directory -Path $BridgeDir -Force|Out-Null

Say "Installed bridge version: $(Get-Version $Server)"
$tmpRoot=Join-Path ([IO.Path]::GetTempPath()) ('clwsc-update-'+[Guid]::NewGuid().ToString('N'))

try{
  New-Item -ItemType Directory -Path $tmpRoot -Force|Out-Null
  Say 'Checking GitHub for the newest bridge runtime...'
  $tree=Invoke-Json $ApiTree
  if([bool]$tree.truncated){throw 'GitHub bridge tree response was truncated; refusing an incomplete update.'}

  $entries=@(
    $tree.tree|
      Where-Object{
        $_.type -eq 'blob' -and
        $_.path -like 'bridge/*' -and
        $_.path -notmatch '^bridge/(config\.json|dxf-index\.json|dxf-index-status\.json)$' -and
        $_.path -notmatch '^bridge/dxf-index/' -and
        $_.path -notmatch '^bridge/backup/'
      }|
      ForEach-Object{[pscustomobject]@{relative=[string]$_.path.Substring(7);sha=[string]$_.sha}}
  )
  if($entries.Count -eq 0){throw 'GitHub returned no bridge runtime files.'}

  foreach($entry in $entries){
    $stage=Join-Path $tmpRoot $entry.relative
    $url=$RawBase+'bridge/'+$entry.relative.Replace('\','/')
    Say "Verifying bridge\$($entry.relative)..."
    Download-Verified $url $stage $entry.sha
    $ext=[IO.Path]::GetExtension($entry.relative).ToLowerInvariant()
    if($ext -eq '.ps1' -and -not(Test-Syntax $stage)){throw "Downloaded $($entry.relative) failed PowerShell syntax validation."}
  }

  $remoteServer=Join-Path $tmpRoot 'server.ps1'
  if(-not(Test-Path -LiteralPath $remoteServer)){throw 'GitHub bridge does not contain server.ps1.'}
  $remoteVersion=Get-Version $remoteServer
  if($remoteVersion -eq 'unknown' -or $remoteVersion -eq 'none'){throw 'GitHub server.ps1 does not publish a bridge version.'}
  Say "Latest bridge version on GitHub: $remoteVersion" 'Cyan'

  $install=@()
  foreach($entry in $entries){
    $local=Join-Path $BridgeDir $entry.relative
    $stage=Join-Path $tmpRoot $entry.relative
    if((Get-Hash $stage) -ne (Get-Hash $local)){
      $install+=[pscustomobject]@{Stage=$stage;Local=$local;Relative=$entry.relative}
    }
  }

  $localManifest=Join-Path $Root 'manifest.xml'
  $manifestInstall=$false
  if(Test-Path -LiteralPath $localManifest){
    $remoteManifest=Join-Path $tmpRoot 'manifest.xml'
    Say 'Checking local manifest.xml...'
    Invoke-WebRequest -Uri ($RawBase+'manifest.xml') -OutFile $remoteManifest -UseBasicParsing -TimeoutSec 120 -Headers @{'User-Agent'='CL-WS-Creator-Updater';'Cache-Control'='no-cache'}
    if(-not(Test-Xml $remoteManifest)){throw 'Downloaded manifest.xml failed XML validation.'}
    if((Get-Hash $remoteManifest) -ne (Get-Hash $localManifest)){$manifestInstall=$true}
  }

  Stop-CreatorProcesses -ServerPath $Server

  $changed=@($install)
  if($manifestInstall){$changed+=[pscustomobject]@{Stage=$remoteManifest;Local=$localManifest;Relative='manifest.xml'}}

  if($changed.Count -gt 0){
    if(Test-Path -LiteralPath $BackupDir){Remove-Item -LiteralPath $BackupDir -Recurse -Force}
    New-Item -ItemType Directory -Path $BackupDir -Force|Out-Null

    $newFiles=@()
    $backups=@()
    try{
      foreach($item in $changed){
        $parent=Split-Path -Parent $item.Local
        if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
        if(Test-Path -LiteralPath $item.Local){
          $backup=Join-Path $BackupDir $item.Relative
          New-Item -ItemType Directory -Path (Split-Path -Parent $backup) -Force|Out-Null
          Copy-Item -LiteralPath $item.Local -Destination $backup -Force
          $backups+=[pscustomobject]@{Local=$item.Local;Backup=$backup}
        }else{$newFiles+=$item.Local}
      }
      foreach($item in $changed){
        Copy-Item -LiteralPath $item.Stage -Destination $item.Local -Force
      }

      foreach($item in $changed){
        $ext=[IO.Path]::GetExtension($item.Local).ToLowerInvariant()
        if($ext -eq '.ps1' -and -not(Test-Syntax $item.Local)){throw "Installed $($item.Relative) failed PowerShell syntax validation."}
        if([IO.Path]::GetFileName($item.Local) -eq 'manifest.xml' -and -not(Test-Xml $item.Local)){throw 'Installed manifest.xml failed XML validation.'}
      }
      Say ('Updated: '+(($changed|ForEach-Object{$_.Relative})-join ', ')) 'Green'
    }catch{
      Say 'Update failed; restoring the previous installed files...' 'Red'
      foreach($item in $backups){Copy-Item -LiteralPath $item.Backup -Destination $item.Local -Force}
      foreach($local in $newFiles){Remove-Item -LiteralPath $local -Force -ErrorAction SilentlyContinue}
      throw
    }
  }else{
    Say 'Bridge files are already current.' 'Green'
  }
}catch{
  Say "Could not update ($($_.Exception.Message)). Using the installed bridge copy." 'Yellow'
}finally{
  Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}

if(-not(Test-Path -LiteralPath $Server)){
  Say 'No working Creator bridge is installed.' 'Red'
  exit 1
}

Say "Starting bridge $(Get-Version $Server) on http://127.0.0.1:$Port" 'Cyan'
$ps=(Join-Path $PSHOME 'powershell.exe')
if(-not(Test-Path -LiteralPath $ps)){ $ps=(Get-Command powershell.exe -ErrorAction Stop).Source }
& $ps -NoProfile -ExecutionPolicy Bypass -File $Server
exit $LASTEXITCODE
