param(
  [Parameter(Mandatory=$true)][string]$Root
)

$ErrorActionPreference='Stop'
$Repo='eithanbat-bot/CL---WS-Creator'
$Branch='main'
$ApiBase="https://api.github.com/repos/$Repo/contents"
$RawBase="https://raw.githubusercontent.com/$Repo/$Branch"
$BridgeDir=Join-Path $Root 'bridge'
$LocalServer=Join-Path $BridgeDir 'server.ps1'
$ManifestPath=Join-Path $Root 'manifest.xml'

function Say([string]$s){ Write-Host "[CL-WS] $s" }

function Download-Json([string]$url){
  $headers=@{'User-Agent'='CL-WS-Creator-Updater';'Accept'='application/vnd.github+json'}
  Invoke-RestMethod -Uri $url -Headers $headers -UseBasicParsing
}

function Download-Raw([string]$url,[string]$dest){
  $tmp=$dest+'.update'
  try{
    Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing -Headers @{'User-Agent'='CL-WS-Creator-Updater'} -TimeoutSec 60
    if(-not(Test-Path -LiteralPath $tmp)){throw 'Download did not create a file.'}
    Move-Item -LiteralPath $tmp -Destination $dest -Force
  }finally{
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  }
}

function Stop-LocalBridge(){
  $needle=$LocalServer.ToLowerInvariant()
  try{
    $procs=Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
    foreach($p in @($procs)){
      $cmd=[string]$p.CommandLine
      if($cmd -and $cmd.ToLowerInvariant().Contains($needle)){
        Say "Stopping previous Creator bridge process (PID $($p.ProcessId))..."
        Stop-Process -Id ([int]$p.ProcessId) -Force -ErrorAction SilentlyContinue
      }
    }
  }catch{
    Say 'Could not inspect previous bridge processes; continuing.'
  }
  Start-Sleep -Milliseconds 300
}

try{
  if(-not(Test-Path -LiteralPath $BridgeDir)){New-Item -ItemType Directory -Path $BridgeDir -Force|Out-Null}
  Stop-LocalBridge
  Say 'Checking GitHub for the newest Creator bridge files...'

  $items=@(Download-Json ($ApiBase + '/bridge?ref=' + $Branch))
  $runtime=@($items|Where-Object {
    $_.type -eq 'file' -and (
      $_.name -match '\.ps1$' -or
      $_.name -match '\.js$'
    ) -and $_.name -notmatch '(^|/)config\.json$'
  })

  if(-not $runtime.Count){throw 'GitHub returned no Creator bridge runtime files.'}

  foreach($item in $runtime){
    $dest=Join-Path $BridgeDir $item.name
    Say "Updating bridge\$($item.name)..."
    Download-Raw ([string]$item.download_url) $dest
  }

  if(Test-Path -LiteralPath $ManifestPath){
    Say 'Updating local Creator manifest.xml...'
    Download-Raw ($RawBase + '/manifest.xml') $ManifestPath
  }

  Say 'Update complete. Local generated DXF index/config data was preserved.'
}catch{
  Say "UPDATE WARNING: $($_.Exception.Message)"
  Say 'Starting the existing local bridge as a fallback.'
}

if(-not(Test-Path -LiteralPath $LocalServer)){
  throw "Creator bridge is missing: $LocalServer"
}

Say 'Starting the Creator PowerShell bridge...'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $LocalServer
