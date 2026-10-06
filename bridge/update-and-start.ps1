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

function Get-Hash([string]$Path){
  if(Test-Path -LiteralPath $Path){return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash}
  return ''
}

function Get-ListeningPids([int]$Port){
  $pids=@()
  try{
    $rows=Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop
    foreach($row in @($rows)){
      $id=[int]$row.OwningProcess
      if($id -gt 0 -and $id -ne $PID){$pids+=$id}
    }
  }catch{
    try{
      $lines=netstat -ano -p tcp 2>$null
      foreach($line in $lines){
        $parts=$line -split '[ ]+' | Where-Object { $_ -ne '' }
        if($parts.Count -ge 5 -and $parts[0] -eq 'TCP' -and $parts[3] -eq 'LISTENING'){
          $local=$parts[1]
          $colon=$local.LastIndexOf(':')
          if($colon -gt 0){
            $localPort=[int]$local.Substring($colon+1)
            if($localPort -eq $Port){
              $id=[int]$parts[4]
              if($id -gt 0 -and $id -ne $PID){$pids+=$id}
            }
          }
        }
      }
    }catch{}
  }
  @($pids|Sort-Object -Unique)
}

function Test-CreatorProcessOwner([int]$ProcessId,[string]$ServerPath){
  try{
    $p=Get-CimInstance Win32_Process -Filter ("ProcessId="+$ProcessId) -ErrorAction Stop
    if($null -eq $p){return $false}
    $cmd=[string]$p.CommandLine
    if([string]::IsNullOrWhiteSpace($cmd)){return $false}
    $full=[IO.Path]::GetFullPath($ServerPath)
    $normCmd=($cmd.Trim('"') -replace '/','\\')
    $normFull=($full -replace '/','\\').TrimEnd('\\')
    $exact=$normCmd.IndexOf($normFull,[StringComparison]::OrdinalIgnoreCase) -ge 0
    $creatorServer=($normCmd -match '(?i)CL---WS-Creator') -and ($normCmd -match '(?i)bridge[\\/]server\.ps1')
    return ($exact -or $creatorServer)
  }catch{return $false}
}

function Stop-CreatorPortOwners([int]$Port,[string]$ServerPath){
  $owners=@(Get-ListeningPids -Port $Port)
  foreach($id in $owners){
    if(-not(Test-CreatorProcessOwner -ProcessId $id -ServerPath $ServerPath)){
      throw ("Creator port "+$Port+" is already owned by PID "+$id+" which does not appear to be the CL-WS Creator bridge. Refusing to terminate an unrelated process.")
    }
    Say "Stopping existing CL-WS Creator bridge PID $id on port $Port..." 'Yellow'
    $stopped=$false
    try{
      Stop-Process -Id $id -Force -ErrorAction Stop
      $stopped=$true
    }catch{
      Say "PowerShell stop failed for PID $id; using taskkill /F..." 'Yellow'
      try{
        & taskkill.exe /PID $id /F /T >$null 2>&1
        if($LASTEXITCODE -eq 0){$stopped=$true}
      }catch{}
    }
    if(-not $stopped){
      throw ("Could not terminate the existing CL-WS Creator bridge PID "+$id+".")
    }
  }

  for($i=0;$i -lt 40;$i++){
    $remaining=@(Get-ListeningPids -Port $Port)
    if($remaining.Count -eq 0){
      Start-Sleep -Milliseconds 250
      $remaining=@(Get-ListeningPids -Port $Port)
      if($remaining.Count -eq 0){return $true}
    }
    foreach($id in $remaining){
      if(Test-CreatorProcessOwner -ProcessId $id -ServerPath $ServerPath){
        try{& taskkill.exe /PID $id /F /T >$null 2>&1}catch{}
      }else{
        throw ("Creator port "+$Port+" remains owned by unrelated PID(s): "+($remaining -join ', '))
      }
    }
    Start-Sleep -Milliseconds 250
  }
  $remaining=@(Get-ListeningPids -Port $Port)
  if($remaining.Count -gt 0){
    throw ("Creator port "+$Port+" is still occupied by PID(s) "+($remaining -join ', ')+" after termination attempts.")
  }
  return $true
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
      ForEach-Object{
        Say "Stopping existing Creator process PID $($_.ProcessId)..."
        Stop-Process -Id ([int]$_.ProcessId) -Force -ErrorAction SilentlyContinue
      }
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
$Launcher=Join-Path $Root 'Start Bridge Fixed.bat'
function Update-Launcher([string]$Destination){
  try{
    $urls=@(
      "https://raw.githubusercontent.com/$Repo/$Branch/Start%20Bridge%20Fixed.bat",
      "https://github.com/$Repo/raw/refs/heads/$Branch/Start%20Bridge%20Fixed.bat"
    )
    $tmp=Join-Path ([IO.Path]::GetTempPath()) ('clwsc-launcher-'+[Guid]::NewGuid().ToString('N')+'.bat')
    foreach($u in $urls){
      try{
        Invoke-WebRequest -UseBasicParsing -TimeoutSec 60 -Uri $u -OutFile $tmp -Headers @{
          'User-Agent'='CL-WS-Creator-Updater'
          'Cache-Control'='no-cache'
        }
        if((Test-Path -LiteralPath $tmp) -and (Get-Item -LiteralPath $tmp).Length -gt 100){
          $newText=Get-Content -LiteralPath $tmp -Raw -Encoding UTF8
          if(-not $newText.Contains('CL - WS Creator - Bridge')){Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue;continue}
          if(-not(Test-Path -LiteralPath $Destination) -or (Get-Hash $tmp) -ne (Get-Hash $Destination)){
            Copy-Item -LiteralPath $tmp -Destination $Destination -Force
            Say 'Updated local Start Bridge Fixed.bat launcher.' 'Green'
          }
          Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
          return
        }
      }catch{}
    }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  }catch{}
}

Say "Installed bridge version: $(Get-Version $Server)"
$tmpRoot=Join-Path ([IO.Path]::GetTempPath()) ('clwsc-update-'+[Guid]::NewGuid().ToString('N'))

try{
  New-Item -ItemType Directory -Path $tmpRoot -Force|Out-Null
  Say 'Checking GitHub for the newest bridge runtime...'
    # Download one repository archive instead of making a separate API request
  # for every bridge file. This avoids GitHub API rate limits/403 responses.
  $zip=Join-Path $tmpRoot 'repository.zip'
  $extract=Join-Path $tmpRoot 'repository'
  $archiveUrl="https://github.com/$Repo/archive/refs/heads/$Branch.zip"
  Say 'Downloading the current repository archive...'
  $archiveUrls=@(
    $archiveUrl,
    "https://codeload.github.com/$Repo/zip/refs/heads/$Branch"
  )
  $downloaded=$false
  $archiveErrors=@()
  foreach($candidateUrl in $archiveUrls){
    try{
      Invoke-WebRequest -Uri $candidateUrl -OutFile $zip -UseBasicParsing -TimeoutSec 180 -Headers @{
        'User-Agent'='CL-WS-Creator-Updater'
        'Cache-Control'='no-cache'
      }
      if((Test-Path -LiteralPath $zip) -and (Get-Item -LiteralPath $zip).Length -gt 100){
        $downloaded=$true
        break
      }
    }catch{
      $archiveErrors+=($candidateUrl+': '+$_.Exception.Message)
    }
  }
  if(-not $downloaded){
    throw ('Could not download the GitHub repository archive. '+($archiveErrors -join ' | '))
  }
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  Expand-Archive -LiteralPath $zip -DestinationPath $extract -Force
  $archiveRoot=@(Get-ChildItem -LiteralPath $extract -Directory -ErrorAction Stop|Select-Object -First 1)
  if($archiveRoot.Count -ne 1){throw 'Could not determine the extracted GitHub repository folder.'}

  $remoteBridge=Join-Path $archiveRoot[0].FullName 'bridge'
  if(-not(Test-Path -LiteralPath $remoteBridge)){throw 'GitHub archive does not contain the bridge folder.'}

  $runtimeFiles=@(Get-ChildItem -LiteralPath $remoteBridge -File -ErrorAction Stop |
    Where-Object{
      $_.Name -notin @('config.json','dxf-index.json','dxf-index-status.json') -and
      $_.Name -notlike '*.log'
    })
  if($runtimeFiles.Count -eq 0){throw 'GitHub archive contains no bridge runtime files.'}

  $entries=@($runtimeFiles|ForEach-Object{
    [pscustomobject]@{
      relative=$_.Name
      source=$_.FullName
    }
  })

  foreach($entry in $entries){
    $validationStage=Join-Path $tmpRoot ('validate-'+$entry.relative)
    Copy-Item -LiteralPath $entry.source -Destination $validationStage -Force
    $ext=[IO.Path]::GetExtension($entry.relative).ToLowerInvariant()
    Say "Verifying bridge\\$($entry.relative)..."
    if(-not(Test-Path -LiteralPath $validationStage)){throw "Downloaded bridge file did not materialize for validation: $($entry.relative)"}
    if($ext -eq '.ps1' -and -not(Test-Syntax $validationStage)){throw "Downloaded $($entry.relative) failed PowerShell syntax validation."}
    if($ext -eq '.xml' -and -not(Test-Xml $validationStage)){throw "Downloaded $($entry.relative) failed XML validation."}
  }

  $remoteServer=Join-Path $remoteBridge 'server.ps1'
  if(-not(Test-Path -LiteralPath $remoteServer)){throw 'GitHub bridge does not contain server.ps1.'}
  $remoteVersion=Get-Version $remoteServer
  if($remoteVersion -eq 'unknown' -or $remoteVersion -eq 'none'){throw 'GitHub server.ps1 does not publish a bridge version.'}
  Say "Latest bridge version on GitHub: $remoteVersion" 'Cyan'

  $install=@()
  foreach($entry in $entries){
    $local=Join-Path $BridgeDir $entry.relative
    $stage=[string]$entry.source
    if((Get-Hash $stage) -ne (Get-Hash $local)){
      $install+=[pscustomobject]@{Stage=[string]$entry.source;Local=$local;Relative=$entry.relative}
    }
  }

  $localManifest=Join-Path $Root 'manifest.xml'
  $manifestInstall=$false
  if(Test-Path -LiteralPath $localManifest){
    $remoteManifest=Join-Path $archiveRoot[0].FullName 'manifest.xml'
    Say 'Checking local manifest.xml from the downloaded archive...'
    if(-not(Test-Path -LiteralPath $remoteManifest)){throw 'GitHub archive does not contain manifest.xml.'}
    $buildId=($remoteVersion -replace '[^A-Za-z0-9]','')
    $manifestText=Get-Content -LiteralPath $remoteManifest -Raw -Encoding UTF8
    if($manifestText.Contains('__BUILD__')){
      $manifestText=$manifestText.Replace('__BUILD__',$buildId)
      Set-Content -LiteralPath $remoteManifest -Value $manifestText -Encoding UTF8
    }
    if(-not(Test-Xml $remoteManifest)){throw 'Downloaded manifest.xml failed XML validation.'}
    if((Get-Hash $remoteManifest) -ne (Get-Hash $localManifest)){$manifestInstall=$true}
  }

  Stop-CreatorProcesses -ServerPath $Server
  Stop-CreatorPortOwners -Port $Port -ServerPath $Server | Out-Null

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
        if(-not(Test-Path -LiteralPath $item.Stage)){throw "Downloaded archive source is missing before install: $($item.Relative)"}
        Copy-Item -LiteralPath $item.Stage -Destination $item.Local -Force
      }

      foreach($item in $changed){
        $ext=[IO.Path]::GetExtension($item.Local).ToLowerInvariant()
        if($ext -eq '.ps1' -and -not(Test-Syntax $item.Local)){throw "Installed $($item.Relative) failed PowerShell syntax validation."}
        if([IO.Path]::GetFileName($item.Local) -eq 'manifest.xml' -and -not(Test-Xml $item.Local)){throw 'Installed manifest.xml failed XML validation.'}
      }
      Say ('Updated: '+(($changed|ForEach-Object{$_.Relative})-join ', ')) 'Green'
      $installedVersion=Get-Version $Server
      if($installedVersion -ne $remoteVersion){throw "Installed bridge version verification failed. Expected $remoteVersion but found $installedVersion."}
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

Update-Launcher -Destination $Launcher

if(-not(Test-Path -LiteralPath $Server)){
  Say 'No working Creator bridge is installed.' 'Red'
  exit 1
}

$finalOwners=@(Get-ListeningPids -Port $Port)
if($finalOwners.Count -gt 0){
  $details=@()
  foreach($id in $finalOwners){
    try{
      $p=Get-CimInstance Win32_Process -Filter ("ProcessId="+$id) -ErrorAction Stop
      $details+=("PID "+$id+": "+([string]$p.CommandLine))
    }catch{
      $details+=("PID "+$id)
    }
  }
  throw ("Bridge startup aborted because port "+$Port+" is still occupied. "+($details -join ' | '))
}

$installedVersion=Get-Version $Server
  Say "Starting bridge $installedVersion on http://127.0.0.1:$Port" 'Cyan'
$ps=(Join-Path $PSHOME 'powershell.exe')
if(-not(Test-Path -LiteralPath $ps)){ $ps=(Get-Command powershell.exe -ErrorAction Stop).Source }
& $ps -NoProfile -ExecutionPolicy Bypass -STA -File $Server
exit $LASTEXITCODE
