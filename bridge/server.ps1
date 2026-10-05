param()

$ErrorActionPreference = 'Stop'
$PORT = 17832
$BRIDGE_VERSION = '2.10.0'
$DEFAULT_LIBRARY = if($env:SN_PARTS){$env:SN_PARTS}else{'S:\SNDataX1\PARTS'}
$DEFAULT_DXF_LIBRARY = if($env:SN_DXF){$env:SN_DXF}else{'Y:\'}
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
$BridgeDir = $ROOT
$CFG_FILE = Join-Path $ROOT 'config.json'
$DXF_INDEX_FILE = Join-Path $ROOT 'dxf-index.json'
$DXF_INDEX_DIR = Join-Path $ROOT 'dxf-index'
$DXF_STATUS_FILE = Join-Path $ROOT 'dxf-index-status.json'
$DXF_SCAN_SCRIPT = Join-Path $ROOT 'dxf-indexer.ps1'
$SN_COM_LIBRARY = Join-Path $ROOT 'sigmanest-com.ps1'
if(-not(Test-Path -LiteralPath $SN_COM_LIBRARY)){
  throw ('SigmaNEST COM library is missing: '+$SN_COM_LIBRARY)
}
. $SN_COM_LIBRARY
$CFG = [ordered]@{ libraryRoot=$DEFAULT_LIBRARY; dxfRoot=$DEFAULT_DXF_LIBRARY; lastScan=$null; count=0; discoveredFiles=0; prsCount=0; dxfCount=0; dxfIndexState='IDLE'; dxfIndexMessage=''; scanErrors=@(); inspectErrors=@(); dxfWorkers=12; dxfNightlyHour=2; dxfRefreshHours=24; dxfAutoRefresh=$true }

try {
  if(Test-Path -LiteralPath $CFG_FILE){
    $old = Get-Content -LiteralPath $CFG_FILE -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach($p in $old.PSObject.Properties){ $CFG[$p.Name] = $p.Value }
  }
} catch {}

try {
  # DXF indexing is intentionally configured at the bridge maximum.
  $CFG['dxfWorkers']=12
  if($null -eq $CFG['dxfNightlyHour']){$CFG['dxfNightlyHour']=2}
  if($null -eq $CFG['dxfRefreshHours']){$CFG['dxfRefreshHours']=24}
  if($null -eq $CFG['dxfAutoRefresh']){$CFG['dxfAutoRefresh']=$true}
} catch {}

function Send-Json($client, [int]$status, $obj) {
  $json = ($obj | ConvertTo-Json -Depth 20 -Compress)
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  $reason = switch($status){200{'OK'};204{'No Content'};400{'Bad Request'};404{'Not Found'};500{'Internal Server Error'};default{'OK'}}
  $crlf=[char]13+[char]10
  $head = "HTTP/1.1 $status $reason$crlf" +
    "Content-Type: application/json; charset=utf-8$crlf" +
    "Access-Control-Allow-Origin: *$crlf" +
    "Access-Control-Allow-Headers: Content-Type, Accept, X-Requested-With$crlf" +
    "Access-Control-Allow-Methods: GET, POST, OPTIONS$crlf" +
    "Access-Control-Allow-Private-Network: true$crlf" +
    "Cache-Control: no-store$crlf" +
    "Content-Length: $($bytes.Length)$crlf" +
    "Connection: close$crlf$crlf"
  $hb = [Text.Encoding]::ASCII.GetBytes($head)
  $stream = $client.GetStream()
  if($hb.Length){$stream.Write($hb,0,$hb.Length)}
  if($bytes.Length){$stream.Write($bytes,0,$bytes.Length)}
  $stream.Flush()
  $stream.Close()
  $client.Close()
}

function Read-HttpRequest($client) {
  $stream = $client.GetStream()
  $ms = New-Object IO.MemoryStream
  $buf = New-Object byte[] 8192
  $headerEnd = -1
  $contentLength = 0
  while($true){
    $n = $stream.Read($buf,0,$buf.Length)
    if($n -le 0){break}
    $ms.Write($buf,0,$n)
    $raw = [Text.Encoding]::ASCII.GetString($ms.ToArray())
    $headerEnd = $raw.IndexOf(([string][char]13+[char]10+[char]13+[char]10))
    if($headerEnd -ge 0){
      $headers = $raw.Substring(0,$headerEnd)
      $m = [regex]::Match($headers,'(?im)^Content-Length:\s*(\d+)\s*$')
      if($m.Success){$contentLength=[int]$m.Groups[1].Value}
      $bodyStart = $headerEnd + 4
      $currentBodyBytes = $ms.Length - $bodyStart
      if($currentBodyBytes -ge $contentLength){break}
    }
    if($ms.Length -gt 10485760){throw 'HTTP request is too large.'}
  }
  if($headerEnd -lt 0){throw 'Invalid HTTP request.'}
  $all = $ms.ToArray()
  $headerText = [Text.Encoding]::ASCII.GetString($all,0,$headerEnd)
  $lines = $headerText -split ([string][char]13+[char]10)
  $requestLine = $lines[0] -split ' '
  $headers = @{}
  if($lines.Length -gt 1){
    foreach($line in $lines[1..($lines.Length-1)]){
      $i=$line.IndexOf(':')
      if($i -gt 0){$headers[$line.Substring(0,$i).Trim().ToLowerInvariant()]=$line.Substring($i+1).Trim()}
    }
  }
  $bodyStart = $headerEnd + 4
  $body = ''
  if($contentLength -gt 0 -and ($all.Length-$bodyStart) -ge $contentLength){
    $body = [Text.Encoding]::UTF8.GetString($all,$bodyStart,$contentLength)
  }
  [pscustomobject]@{ Method=$requestLine[0]; Path=($requestLine[1] -split '\?')[0]; Body=$body; Headers=$headers }
}

function Get-Strings8([byte[]]$b) {
  $s = [Text.Encoding]::ASCII.GetString($b)
  [regex]::Matches($s,'[\x20-\x7E]{4,}') | ForEach-Object {$_.Value}
}
function Get-Strings16([byte[]]$b) {
  $s = [Text.Encoding]::Unicode.GetString($b)
  [regex]::Matches($s,'[\x20-\x7E]{4,}') | ForEach-Object {$_.Value}
}
function Set-Prop($obj,[string]$name,$value){
  if([string]::IsNullOrWhiteSpace($name)){return}
  try{ $obj | Add-Member -MemberType NoteProperty -Name $name -Value $value -Force }catch{}
  return $obj
}

function Inspect-Prs([string]$file) {
  $b=[IO.File]::ReadAllBytes($file)
  $strings = @((Get-Strings8 $b)+(Get-Strings16 $b) | Where-Object {$_} | Select-Object -Unique)
  $stem=[IO.Path]::GetFileNameWithoutExtension($file)
  $embedded=($strings | Where-Object {$_ -like "$stem*"} | Select-Object -First 1)
  if(-not $embedded){$embedded=$stem}
  $source=($strings | Where-Object {$_ -match '\.(dxf|dwg)$'} | Select-Object -First 1)
  if(-not $source){$source=''}
  $material=$strings | Where-Object {
    $_.Trim() -match '^(armox(?:\s+advance|\s+\d+)?|ramor(?:\s+\d+)?|s\d{3,4}|hardox(?:\s+\d+)?|chromodeck|mild steel|strenx(?:\s+\d+)?|aluminium|aluminum|stainless(?: steel)?)$'
  } | Select-Object -First 1
  if(-not $material){
    $material=$strings | Where-Object {$_.Length -lt 80 -and $_ -match '(?i)armox|ramor|s355|s690|hardox|chromodeck|mild steel|strenx|aluminium|aluminum|stainless'} | Sort-Object Length | Select-Object -First 1
  }
  $m=[regex]::Match([string]$material,'(\d+(?:\.\d+)?)\s*mm')
  if(-not $m.Success){$m=[regex]::Match([string]$embedded,'-\s*(\d+(?:\.\d+)?)\s*mm')}
  $thickness=if($m.Success){$m.Groups[1].Value+'mm'}else{''}
  [pscustomobject]@{
    file=[IO.Path]::GetFullPath($file); fileName=[IO.Path]::GetFileName($file)
    partName=$stem; embeddedPartName=$embedded; likelyMaterial=[string]$material
    thickness=$thickness; sourceDxf=[string]$source; rotations=''
  }
}

function Normalize([string]$s){ if($null -eq $s){$s=''}; (($s.ToUpperInvariant() -replace '[^A-Z0-9]','') -replace 'PRS$','') }
function VariationKey([string]$s){ $n=Normalize -s $s; if($n -match '[A-Z0-9]$'){ $n=$n.Substring(0,$n.Length-1) }; return $n }
function Normalize-Material([string]$s){
  if($null -eq $s){$s=''}
  $s=$s -replace '(?i)Armoxt|Amoxt','Armox'
  $s=$s -replace '(?i)Ramort','Ramor'
  ($s -replace '\s+',' ').Trim()
}
function Thickness-Number([string]$s){
  $m=[regex]::Match([string]$s,'(\d+(?:\.\d+)?)\s*mm',[Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if($m.Success){[double]$m.Groups[1].Value}else{[double]::NaN}
}
function Material-Equal([string]$a,[string]$b,[string]$aThickness='',[string]$bThickness=''){
  $A=(Normalize-Material -s $a).ToUpperInvariant()
  $B=(Normalize-Material -s $b).ToUpperInvariant()
  if(-not $A -or -not $B -or $A -eq $B){return $true}

  # Required job rule: 4 mm Armox is treated as Ramor 500 4 mm.
  $aThk=Thickness-Number -s $aThickness
  $bThk=Thickness-Number -s $bThickness
  $A4Armox=($A -match 'ARMOX' -and (($aThk -eq 4) -or ($A -match '4\s*MM')))
  $B4Armox=($B -match 'ARMOX' -and (($bThk -eq 4) -or ($B -match '4\s*MM')))
  $ARamor500=($A -match 'RAMOR' -and $A -match '500' -and (($aThk -eq 4) -or ($A -match '4\s*MM')))
  $BRamor500=($B -match 'RAMOR' -and $B -match '500' -and (($bThk -eq 4) -or ($B -match '4\s*MM')))
  if(($A4Armox -and $BRamor500) -or ($B4Armox -and $ARamor500)){return $true}

  $ANoSize=($A -replace '\b\d+(?:\.\d+)?\s*MM\b','' -replace '\s+(SHEET|PLATE)\b','').Trim()
  $BNoSize=($B -replace '\b\d+(?:\.\d+)?\s*MM\b','' -replace '\s+(SHEET|PLATE)\b','').Trim()
  if($ANoSize -eq $BNoSize){return $true}
  return ($ANoSize.Contains($BNoSize) -or $BNoSize.Contains($ANoSize))
}

function Sigma-Material([string]$cl,[string]$lib){
  $s=Normalize-Material -s ([string]$(if($lib){$lib}else{$cl}))
  $s=$s -replace '^\d+(?:\.\d+)?\s*mm\s*',''
  $s=$s -replace '\s+(sheet|plate)$',''
  $s.Trim()
}
function Get-DxfStatus(){
  if(-not(Test-Path -LiteralPath $DXF_STATUS_FILE)){return [pscustomobject]@{state='IDLE';root=$DEFAULT_DXF_LIBRARY;filesFound=0;errors=0;message='DXF index has not been started.'}}
  try{return (Get-Content -LiteralPath $DXF_STATUS_FILE -Raw -Encoding UTF8|ConvertFrom-Json)}catch{return [pscustomobject]@{state='UNKNOWN';root=$DEFAULT_DXF_LIBRARY;filesFound=0;errors=0;message='Could not read DXF index status.'}}
}

function Get-PowerShellExe(){
  $candidate=Join-Path $PSHOME 'powershell.exe'
  if(Test-Path -LiteralPath $candidate){return $candidate}
  return (Get-Command powershell.exe -ErrorAction Stop).Source
}

function Stop-RunningDxfWorkers([string]$root){
  try{
    $rootPattern=[regex]::Escape($root)
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
      Where-Object {
        $_.ProcessId -ne $PID -and
        ([string]$_.CommandLine -match 'dxf-indexer\.ps1')
      } |
      ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
      }
  }catch{}
}

function Start-DxfScan([string]$root,[string]$mode='FULL',[switch]$force){
  $status=Get-DxfStatus
  $requiredIndexerVersion='3.0.1'
  $rootFull=[IO.Path]::GetFullPath(([string]$root).Trim())
  $sameRoot=([string]$status.root).Equals($rootFull,[StringComparison]::OrdinalIgnoreCase)

  if($sameRoot -and [string]$status.state -eq 'RUNNING'){
    try{
      if($status.pid){
        $p=Get-Process -Id ([int]$status.pid) -ErrorAction SilentlyContinue
        if($p){return $status}
      }
    }catch{}
  }

  if($sameRoot -and [string]$status.state -eq 'COMPLETE' -and
     [string]$status.indexerVersion -eq $requiredIndexerVersion -and
     (Test-Path -LiteralPath $DXF_INDEX_FILE) -and
     (Test-Path -LiteralPath $DXF_INDEX_DIR) -and
     -not $force){
    return $status
  }

  if(-not $sameRoot -or [string]$status.state -notin @('RUNNING','COMPLETE') -or [string]$status.indexerVersion -ne $requiredIndexerVersion){
    $mode='FULL'
  } elseif($mode -ne 'REFRESH'){
    $mode='FULL'
  }

  if(-not(Test-Path -LiteralPath $DXF_SCAN_SCRIPT)){
    return [pscustomobject]@{
      state='FAILED'; indexerVersion=$requiredIndexerVersion; root=$rootFull; filesFound=0; errors=1
      message='DXF indexer script is missing: '+$DXF_SCAN_SCRIPT
    }
  }

  if([string]$status.state -eq 'RUNNING'){
    return $status
  }

  try{
    $started=(Get-Date).ToUniversalTime().ToString('o')
    $workerCount=[int]$CFG['dxfWorkers']
    if($workerCount -le 0){$workerCount=12}
    $workerCount=[math]::Max(1,[math]::Min(12,$workerCount))

    $starting=[pscustomobject]@{
      state='RUNNING'
      indexerVersion=$requiredIndexerVersion
      root=$rootFull
      mode=$mode
      message=($mode+' DXF index controller is starting in the background with '+$workerCount+' worker(s).')
      filesFound=$(try{[int]$status.filesFound}catch{0})
      errors=0
      started=$started
      finished=$null
      generatedUtc=$(try{[string]$status.generatedUtc}catch{''})
      currentPath=$rootFull
      pid=$null
      workers=$workerCount
      workersCompleted=0
      directoriesVisited=$(try{[int]$status.directoriesVisited}catch{0})
      elapsedSeconds=0
      logFile=(Join-Path $ROOT 'dxf-indexer.log')
      workDirectory=''
    }
    ($starting|ConvertTo-Json -Depth 12)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8

    $requestFile=Join-Path $ROOT ('_dxf-index-request-'+[Guid]::NewGuid().ToString('N')+'.json')
    $request=[ordered]@{
      Role='CONTROLLER'
      Mode=$mode
      WorkerCount=$workerCount
      Root=$rootFull
      IndexFile=$DXF_INDEX_FILE
      StatusFile=$DXF_STATUS_FILE
      CreatedUtc=(Get-Date).ToUniversalTime().ToString('o')
    }
    ($request|ConvertTo-Json -Depth 12)|Set-Content -LiteralPath $requestFile -Encoding UTF8

    $launchLog=Join-Path $ROOT 'dxf-indexer-launch.log'
    $psExe=Get-PowerShellExe
    $psi=New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName=$psExe
    $psi.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$DXF_SCAN_SCRIPT+'" -Role CONTROLLER -RequestFile "'+$requestFile+'"'
    $psi.WorkingDirectory=$ROOT
    $psi.UseShellExecute=$false
    $psi.CreateNoWindow=$true
    $psi.RedirectStandardOutput=$false
    $psi.RedirectStandardError=$false
    $proc=[System.Diagnostics.Process]::Start($psi)
    $starting.pid=$proc.Id
    ($starting|ConvertTo-Json -Depth 12)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8
    ('['+(Get-Date).ToUniversalTime().ToString('o')+'] '+$mode+' controller launched PID='+$proc.Id+' workers='+$workerCount)|Add-Content -LiteralPath $launchLog -Encoding UTF8

    return $starting
  }catch{
    Remove-Item -LiteralPath $requestFile -Force -ErrorAction SilentlyContinue
    $message='Could not launch DXF index controller: '+$_.Exception.Message
    $failed=[pscustomobject]@{
      state='FAILED'; indexerVersion=$requiredIndexerVersion; root=$rootFull; filesFound=0; errors=1
      message=$message; pid=$null; logFile=(Join-Path $ROOT 'dxf-indexer.log'); launchLogFile=(Join-Path $ROOT 'dxf-indexer-launch.log')
    }
    try{($failed|ConvertTo-Json -Depth 12)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8}catch{}
    return $failed
  }
}
function Dxf-ShardKey([string]$part){
  $n=Normalize -s $part
  if([string]::IsNullOrWhiteSpace($n)){return 'S___'}
  if($n.Length -ge 3){return 'S_'+$n.Substring(0,3)}
  while($n.Length -lt 3){$n+='_'}
  return 'S_'+$n
}

function Get-DxfShardCandidates([string]$root,[string]$part){
  if([string]::IsNullOrWhiteSpace($root)){return @()}
  if(-not(Test-Path -LiteralPath $DXF_INDEX_FILE)){return @()}
  if(-not(Test-Path -LiteralPath $DXF_INDEX_DIR)){return @()}

  try{
    $manifest=Get-Content -LiteralPath $DXF_INDEX_FILE -Raw -Encoding UTF8|ConvertFrom-Json
    if([string]$manifest.schema -ne 'cl-ws-creator/dxf-index/3'){return @()}
    if(-not ([string]$manifest.root).Equals([string]$root,[StringComparison]::OrdinalIgnoreCase)){return @()}
    if([string]$manifest.indexerVersion -ne '3.0.1'){return @()}
  }catch{return @()}

  $shard=Join-Path $DXF_INDEX_DIR ((Dxf-ShardKey -part $part)+'.tsv')
  if(-not(Test-Path -LiteralPath $shard)){return @()}

  if($script:DXF_SHARD_CACHE -and $script:DXF_SHARD_CACHE.ContainsKey($shard)){return $script:DXF_SHARD_CACHE[$shard]}

  $items=@()
  $reader=$null
  try{
    $reader=New-Object IO.StreamReader($shard,[Text.Encoding]::UTF8,$true)
    [void]$reader.ReadLine()
    while(($line=$reader.ReadLine()) -ne $null){
      $tab=$line.IndexOf([char]9)
      if($tab -lt 1){continue}
      $partName=$line.Substring(0,$tab)
      $file=$line.Substring($tab+1)
      if([string]::IsNullOrWhiteSpace($file)){continue}
      $items += [pscustomobject]@{
        file=$file
        fileName=[IO.Path]::GetFileName($file)
        partName=$partName
        embeddedPartName=$partName
        fileType='DXF'
        likelyMaterial=''
        thickness=''
        sourceDxf=$file
        rotations=''
        metadataLoaded=$true
      }
    }
  }catch{}finally{
    if($reader){$reader.Dispose()}
  }
  if($script:DXF_SHARD_CACHE -ne $null){$script:DXF_SHARD_CACHE[$shard]=$items}
  return $items
}

function Is-EmbeddedMatch([string]$candidate,[string]$target){
  $a=Normalize -s $candidate
  $b=Normalize -s $target
  return ($a.Length -gt $b.Length -and $a.StartsWith($b,[StringComparison]::Ordinal))
}

function Get-GeometryFiles([string]$root,[string]$extension,[System.Collections.ArrayList]$errors){
  if([string]::IsNullOrWhiteSpace($root)){return @()}
  try{
    if(-not(Test-Path -LiteralPath $root)){throw "Cannot access geometry folder: $root"}
    if(-not((Get-Item -LiteralPath $root).PSIsContainer)){throw "Geometry path is not a folder: $root"}
    @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter ('*'+$extension) -ErrorAction SilentlyContinue|ForEach-Object{$_.FullName})
  }catch{[void]$errors.Add([pscustomobject]@{path=$root;error=$_.Exception.Message});@()}
}

function Scan-Libraries([string]$prsRoot,[string]$dxfRoot){
  $scanErrors=New-Object System.Collections.ArrayList
  # Never treat our own generated job/staging directory as part of the PRS
  # library. It lives under PARTS for convenience, but its copies are not
  # source geometry and can otherwise create self-copy errors on later builds.
  $prsFiles=@(Get-GeometryFiles -root $prsRoot -extension '.prs' -errors $scanErrors |
    Where-Object {
      $full=[IO.Path]::GetFullPath([string]$_)
      $stagingRoot=[IO.Path]::GetFullPath((Join-Path -Path $prsRoot -ChildPath '_CL_WS_BUILDER'))
      -not $full.StartsWith(($stagingRoot.TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)
    })
  $dxfStatus=Get-DxfStatus

  $items=@()
  foreach($file in $prsFiles){
    $items += [pscustomobject]@{
      file=$file
      fileName=[IO.Path]::GetFileName($file)
      partName=[IO.Path]::GetFileNameWithoutExtension($file)
      embeddedPartName=[IO.Path]::GetFileNameWithoutExtension($file)
      fileType='PRS'
      likelyMaterial=''
      thickness=''
      sourceDxf=''
      rotations=''
      metadataLoaded=$false
    }
  }

  # Keep the DXF tree out of memory. The background index is queried by shard
  # during a build, while the relatively small PRS library remains in memory.
  $script:INDEX=$items
  $script:DXF_ROOT=$dxfRoot
  $script:DXF_SHARD_CACHE=@{}
  $script:BYNAME=@{}
  $script:BYVAR=@{}
  foreach($item in $items){
    $nk=Normalize -s $item.partName
    if($nk){
      if(-not $script:BYNAME.ContainsKey($nk)){$script:BYNAME[$nk]=@()}
      $script:BYNAME[$nk]+=$item
    }
    $vk=VariationKey -s $item.partName
    if($vk){
      if(-not $script:BYVAR.ContainsKey($vk)){$script:BYVAR[$vk]=@()}
      $script:BYVAR[$vk]+=$item
    }
  }

  $dxfCount=0
  try{$dxfCount=[int]$dxfStatus.filesFound}catch{}
  $total=$prsFiles.Count+$dxfCount
  $CFG['libraryRoot']=$prsRoot
  $CFG['dxfRoot']=$dxfRoot
  $CFG['lastScan']=(Get-Date).ToUniversalTime().ToString('o')
  $CFG['count']=$total
  $CFG['discoveredFiles']=$total
  $CFG['prsCount']=$prsFiles.Count
  $CFG['dxfCount']=$dxfCount
  $CFG['dxfIndexState']=[string]$dxfStatus.state
  $CFG['dxfIndexMessage']=[string]$dxfStatus.message
  $CFG['scanErrors']=@($scanErrors)
  $CFG['inspectErrors']=@()
  try{($CFG|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath $CFG_FILE -Encoding UTF8}catch{}

  [pscustomobject]@{
    count=$total
    discoveredFiles=$total
    prsCount=$prsFiles.Count
    dxfCount=$dxfCount
    dxfStatus=$dxfStatus
    scanErrors=@($scanErrors)
    inspectErrors=@()
    prsRoot=$prsRoot
    dxfRoot=$dxfRoot
  }
}
function Ensure-Metadata($item) {
  if($item.metadataLoaded){return $item}
  try{
    $meta=Inspect-Prs ([string]$item.file)
    foreach($p in $meta.PSObject.Properties){
      Set-Prop $item ([string]$p.Name) $p.Value | Out-Null
    }
    Set-Prop $item 'metadataLoaded' $true | Out-Null
  }catch{
    $item|Add-Member NoteProperty metadataLoaded $true -Force
    Set-Prop $item 'metadataError' $_.Exception.Message | Out-Null
  }
  return $item
}

function Deduplicate-Candidates($candidates){
  $seen=@{}
  $out=@()
  foreach($c in @($candidates)){
    if(-not $c){continue}
    $key=''
    try{$key=[IO.Path]::GetFullPath([string]$c.file).ToUpperInvariant()}catch{$key=[string]$c.file}
    if([string]::IsNullOrWhiteSpace($key)){continue}
    if(-not $seen.ContainsKey($key)){
      $seen[$key]=$true
      $out+=$c
    }
  }
  return @($out)
}

function Candidate-Score($candidate,[string]$clMat,[string]$clThk,[string]$matchType){
  $score=switch($matchType){
    'EXACT' {300}
    'EMBEDDED' {200}
    'VARIATION' {100}
    default {0}
  }

  if([string]$candidate.fileType -eq 'PRS'){
    Ensure-Metadata -item $candidate | Out-Null
    $libMat=[string]$candidate.likelyMaterial
    $libThk=[string]$candidate.thickness

    $clMatKnown=-not [string]::IsNullOrWhiteSpace($clMat)
    $clThkKnown=-not [string]::IsNullOrWhiteSpace($clThk)
    $libMatKnown=-not [string]::IsNullOrWhiteSpace($libMat)
    $libThkKnown=-not [string]::IsNullOrWhiteSpace($libThk)

    if($clMatKnown -and $libMatKnown){
      if(Material-Equal -a $clMat -b $libMat -aThickness $clThk -bThickness $libThk){$score+=80}
      else{$score-=220}
    }elseif($clMatKnown -and -not $libMatKnown){
      $score+=5
    }

    if($clThkKnown -and $libThkKnown){
      $ct=Thickness-Number -s $clThk
      $lt=Thickness-Number -s $libThk
      if(-not [double]::IsNaN($ct) -and -not [double]::IsNaN($lt)){
        if($ct -eq $lt){$score+=60}else{$score-=160}
      }
    }elseif($clThkKnown -and -not $libThkKnown){
      $score+=3
    }
  }else{
    # DXFs are deliberately not guessed by material/thickness because their
    # index currently contains path/name metadata only.
    $score+=0
  }

  return [int]$score
}

function Select-MatchCandidate($candidates,[string]$clMat='',[string]$clThk='',[string]$matchType='EXACT') {
  $c=@(Deduplicate-Candidates $candidates)
  if($c.Count -eq 0){return $null}

  # Geometry-source rule:
  # 1) PRS is always preferred over DXF.
  # 2) When multiple PRS candidates exist, material/thickness scoring selects
  #    the best compatible PRS instead of arbitrarily taking the first file.
  # 3) If no PRS exists, the first indexed DXF remains the deterministic choice.
  $prs=@($c|Where-Object {[string]$_.fileType -eq 'PRS'})
  if($prs.Count -gt 0){
    $scored=@($prs|ForEach-Object{
      $score=Candidate-Score -candidate $_ -clMat $clMat -clThk $clThk -matchType $matchType
      [pscustomobject]@{candidate=$_;score=$score}
    }|Sort-Object score -Descending)
    if($scored.Count -eq 1){return $scored[0].candidate}
    $top=[int]$scored[0].score
    $second=[int]$scored[1].score
    # A strong material/thickness winner is safe. A tie or weak lead remains
    # a review case rather than silently selecting the wrong geometry.
    if($top -gt $second -and ($top-$second) -ge 40){return $scored[0].candidate}
    # Keep the best geometry even when the score is tied. The caller will
    # mark the row REVIEW, but it must still retain sourcePath so SigmaNEST can
    # import the geometry. The previous implementation returned only an
    # "ambiguous" wrapper, which deliberately discarded the usable source and
    # caused "Imported into WS: 0".
    $selected=$scored[0].candidate
    Set-Prop $selected 'selectionAmbiguous' $true | Out-Null
    Set-Prop $selected 'candidateScores' @($scored|Select-Object -First 6|ForEach-Object{
      [pscustomobject]@{file=$_.candidate.file;score=$_.score}
    }) | Out-Null
    Set-Prop $selected 'ambiguousCount' ([int]$scored.Count) | Out-Null
    return $selected
  }

  $dxf=@($c|Where-Object {[string]$_.fileType -eq 'DXF'})
  if($dxf.Count -gt 0){return $dxf[0]}
  return $c[0]
}

function Find-Part([string]$part,[string]$clMat='',[string]$clThk=''){
  $n=Normalize -s $part
  $dxf=@()
  if($script:DXF_ROOT){$dxf=@(Get-DxfShardCandidates -root $script:DXF_ROOT -part $part)}

  $prsExact=@()
  if($script:BYNAME -and $script:BYNAME.ContainsKey($n)){$prsExact=@($script:BYNAME[$n])}
  $dxfExact=@($dxf|Where-Object {(Normalize -s $_.partName) -eq $n})
  $exact=@(Deduplicate-Candidates ($prsExact+$dxfExact))
  if($exact.Count){
    $hit=Select-MatchCandidate -candidates $exact -clMat $clMat -clThk $clThk -matchType 'EXACT'
    if($hit -and -not($hit.PSObject.Properties.Name -contains 'ambiguous')){
      Ensure-Metadata -item $hit | Out-Null
      Set-Prop $hit 'matchType' 'EXACT' | Out-Null
      return $hit
    }
    if($hit){return $hit}
  }

  $prsEmbedded=@($script:INDEX|Where-Object {Is-EmbeddedMatch -candidate $_.embeddedPartName -target $part})
  $dxfEmbedded=@($dxf|Where-Object {Is-EmbeddedMatch -candidate $_.embeddedPartName -target $part})
  $embedded=@(Deduplicate-Candidates ($prsEmbedded+$dxfEmbedded))
  if($embedded.Count){
    $hit=Select-MatchCandidate -candidates $embedded -clMat $clMat -clThk $clThk -matchType 'EMBEDDED'
    if($hit -and -not($hit.PSObject.Properties.Name -contains 'ambiguous')){
      Ensure-Metadata -item $hit | Out-Null
      Set-Prop $hit 'matchType' 'EMBEDDED' | Out-Null
      return $hit
    }
    if($hit){return $hit}
  }

  $vk=VariationKey -s $part
  $prsVars=@()
  if($script:BYVAR -and $script:BYVAR.ContainsKey($vk)){$prsVars=@($script:BYVAR[$vk])}
  $dxfVars=@($dxf|Where-Object {(VariationKey -s $_.partName) -eq $vk})
  $vars=@(Deduplicate-Candidates ($prsVars+$dxfVars))
  if($vars.Count){
    $hit=Select-MatchCandidate -candidates $vars -clMat $clMat -clThk $clThk -matchType 'VARIATION'
    if($hit){
      Ensure-Metadata -item $hit | Out-Null
      Set-Prop $hit 'matchType' 'VARIATION' | Out-Null
      return $hit
    }
  }

  # Conservative assembly-prefix fallback. This is specifically for the
  # documented case where the CL and library drawing numbers share the
  # assembly/initial characters but differ in a trailing letter/digit.
  $all=@()
  $all += @($script:INDEX)
  $all += @($dxf)
  $all=@(Deduplicate-Candidates $all)
  $target=Normalize -s $part
  if($target.Length -ge 6){
    $fuzzy=@()
    foreach($candidate in $all){
      $cn=Normalize -s ([string]$candidate.partName)
      if($cn.Length -lt 6){continue}
      $maxPrefix=[math]::Min($target.Length,$cn.Length)
      $prefix=0
      while($prefix -lt $maxPrefix -and $target[$prefix] -eq $cn[$prefix]){$prefix++}
      if($prefix -lt 6){continue}
      $minLen=[math]::Min($target.Length,$cn.Length)
      if($prefix -lt [math]::Ceiling($minLen*0.75)){continue}
      $score=$prefix*10 - [math]::Abs($target.Length-$cn.Length)*3
      if([string]$candidate.fileType -eq 'PRS'){
        $score+=20
        Ensure-Metadata -item $candidate | Out-Null
        if(Material-Equal -a $clMat -b ([string]$candidate.likelyMaterial) -aThickness $clThk -bThickness ([string]$candidate.thickness)){$score+=40}
        if((Thickness-Number -s $clThk) -eq (Thickness-Number -s ([string]$candidate.thickness))){$score+=20}
      }
      $fuzzy += [pscustomobject]@{candidate=$candidate;score=$score;prefix=$prefix}
    }
    $fuzzy=@($fuzzy|Sort-Object score -Descending)
    if($fuzzy.Count){
      $top=$fuzzy[0]
      $second=if($fuzzy.Count -gt 1){$fuzzy[1]}else{$null}
      if($null -eq $second -or ([int]$top.score-[int]$second.score) -ge 15){
        $hit=$top.candidate
        Set-Prop $hit 'selectionFuzzy' $true | Out-Null
        Set-Prop $hit 'fuzzyPrefixLength' ([int]$top.prefix) | Out-Null
        Set-Prop $hit 'matchType' 'FUZZY' | Out-Null
        return $hit
      }
    }
  }
  $null
}
function Csv([object]$v){
  $s=[string]$v
  if($s -match '[,"\r\n]'){'"' + ($s -replace '"','""') + '"'}else{$s}
}
function Write-Job([string]$root,[string]$name,$parts){
  $out=Join-Path -Path $root -ChildPath ('_CL_WS_BUILDER\'+$name)
  $partsDir=Join-Path -Path $out -ChildPath 'parts'
  New-Item -ItemType Directory -Path $partsDir -Force|Out-Null

  foreach($p in @($parts|Where-Object {$_.sourcePath})){
    $source=[IO.Path]::GetFullPath([string]$p.sourcePath)
    $dest=Join-Path -Path $partsDir -ChildPath ([IO.Path]::GetFileName($source))
    $dest=[IO.Path]::GetFullPath($dest)
    # A prior generated job may already be the selected source (for example
    # before the server has been restarted with the staging-directory filter).
    # Do not ask Copy-Item to overwrite a file with itself.
    if($source.Equals($dest,[StringComparison]::OrdinalIgnoreCase)){continue}
    Copy-Item -LiteralPath $source -Destination $dest -Force
  }

  $rows=@('Part,Qty,Material,Thickness,SourceType,SourceFile,MatchType,PRS_Material,PRS_Thickness,SourcePath,Status,ReviewReason')
  foreach($p in $parts){
    $vals=@(
      $p.part,$p.qty,$p.material,$p.thickness,$p.sourceType,
      $(if($p.sourcePath){[IO.Path]::GetFileName([string]$p.sourcePath)}else{''}),
      $p.matchType,$p.libraryMaterial,$p.libraryThickness,$p.sourcePath,$p.status,$p.reviewReason
    )
    $rows += (($vals|ForEach-Object{Csv -v $_}) -join ',')
  }
  $rowsFile=Join-Path -Path $out -ChildPath 'WS_PARTS.csv'
  $rows|Set-Content -LiteralPath $rowsFile -Encoding UTF8

  $jsonFile=Join-Path -Path $out -ChildPath 'SIGMANEST_JOB.json'
  ($parts|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath $jsonFile -Encoding UTF8

  $reviewFile=Join-Path -Path $out -ChildPath 'PART_REVIEW.csv'
  $reviewRows=@('Part,Qty,Material,Thickness,Status,ReviewReason,MatchType,SourceType,SourceFile,SourcePath')
  foreach($p in @($parts|Where-Object {$_.status -ne 'READY'})){
    $vals=@($p.part,$p.qty,$p.material,$p.thickness,$p.status,$p.reviewReason,$p.matchType,$p.sourceType,$(if($p.sourcePath){[IO.Path]::GetFileName([string]$p.sourcePath)}else{''}),$p.sourcePath)
    $reviewRows += (($vals|ForEach-Object{Csv -v $_}) -join ',')
  }
  $reviewRows|Set-Content -LiteralPath $reviewFile -Encoding UTF8
  $out
}

function Get-DxfWorkerCount(){
  try{
    $n=[int]$CFG['dxfWorkers']
    if($n -le 0){$n=8}
    return [math]::Max(1,[math]::Min(12,$n))
  }catch{return 8}
}

function Get-DxfRefreshAgeHours($status){
  try{
    $stamp=[DateTime]::Parse([string]$status.finished).ToUniversalTime()
    return ((Get-Date).ToUniversalTime()-$stamp).TotalHours
  }catch{
    return [double]::PositiveInfinity
  }
}

function Get-NextDxfRefreshLocal(){
  try{
    $hour=[int]$CFG['dxfNightlyHour']
    if($hour -lt 0 -or $hour -gt 23){$hour=2}
  }catch{$hour=2}
  $now=Get-Date
  $next=Get-Date -Hour $hour -Minute 0 -Second 0
  if($next -le $now){$next=$next.AddDays(1)}
  $next.ToString('yyyy-MM-dd HH:mm')
}

function Initialize-DxfScheduler(){
  try{
    if($CFG['dxfAutoRefresh'] -ne $true){return}
    $status=Get-DxfStatus
    if([string]$status.state -eq 'RUNNING'){return}
    if(-not(Test-Path -LiteralPath $DXF_INDEX_FILE) -or -not(Test-Path -LiteralPath $DXF_INDEX_DIR)){
      Start-DxfScan -root ([string]$(if($CFG['dxfRoot']){$CFG['dxfRoot']}else{$DEFAULT_DXF_LIBRARY})) -mode 'FULL' | Out-Null
      return
    }
    try{
      $age=Get-DxfRefreshAgeHours -status $status
      $hour=(Get-Date).Hour
      $nightlyHour=[int]$CFG['dxfNightlyHour']
      if($nightlyHour -lt 0 -or $nightlyHour -gt 23){$nightlyHour=2}
      if($age -ge [double]$CFG['dxfRefreshHours'] -and $hour -ge $nightlyHour -and $hour -lt ($nightlyHour+2)){
        Start-DxfScan -root ([string]$(if($CFG['dxfRoot']){$CFG['dxfRoot']}else{$DEFAULT_DXF_LIBRARY})) -mode 'REFRESH' -force | Out-Null
      }
    }catch{}
  }catch{}
}

function Invoke-DxfScheduler(){
  try{
    if($CFG['dxfAutoRefresh'] -ne $true){return}
    $status=Get-DxfStatus
    if([string]$status.state -eq 'RUNNING'){return}
    if(-not(Test-Path -LiteralPath $DXF_INDEX_FILE) -or -not(Test-Path -LiteralPath $DXF_INDEX_DIR)){return}

    $now=Get-Date
    $nightlyHour=[int]$CFG['dxfNightlyHour']
    if($nightlyHour -lt 0 -or $nightlyHour -gt 23){$nightlyHour=2}
    if($now.Hour -lt $nightlyHour -or $now.Hour -ge ($nightlyHour+2)){return}

    $age=Get-DxfRefreshAgeHours -status $status
    $interval=[double]$CFG['dxfRefreshHours']
    if($age -ge $interval){
      Start-DxfScan -root $DEFAULT_DXF_LIBRARY -mode 'REFRESH' -force | Out-Null
    }
  }catch{}
}

function Write-BuildStatus($statusFile,$obj){
  $dir=Split-Path -Parent $statusFile
  New-Item -ItemType Directory -Path $dir -Force|Out-Null
  $tmp=$statusFile+'.tmp-'+[Guid]::NewGuid().ToString('N')
  ($obj|ConvertTo-Json -Depth 30)|Set-Content -LiteralPath $tmp -Encoding UTF8
  Move-Item -LiteralPath $tmp -Destination $statusFile -Force
}

function Get-BuildStatusPath([string]$jobId){
  if([string]::IsNullOrWhiteSpace($jobId)){throw 'Build job ID is required.'}
  if($jobId -notmatch '^[0-9a-fA-F-]{36}$'){throw 'Invalid build job ID.'}
  $dir=Join-Path $BridgeDir 'build-status'
  $path=Join-Path $dir ($jobId+'.json')
  $full=[IO.Path]::GetFullPath($path)
  $rootFull=[IO.Path]::GetFullPath($dir)
  if(-not $full.StartsWith(($rootFull.TrimEnd('')+''),[StringComparison]::OrdinalIgnoreCase)){throw 'Invalid build status path.'}
  return $full
}

function Start-SigmaNestBuildWorker($request){
  $statusDir=Join-Path $BridgeDir 'build-status'
  New-Item -ItemType Directory -Path $statusDir -Force|Out-Null
  $jobId=[Guid]::NewGuid().ToString()
  $statusFile=Get-BuildStatusPath -jobId $jobId
  $requestFile=Join-Path $statusDir ($jobId+'.request.json')
  $request.statusFile=$statusFile
  $request.jobId=$jobId
  ($request|ConvertTo-Json -Depth 30)|Set-Content -LiteralPath $requestFile -Encoding UTF8
  $worker=Join-Path $BridgeDir 'build-worker.ps1'
  if(-not(Test-Path -LiteralPath $worker)){
    Remove-Item -LiteralPath $requestFile -Force -ErrorAction SilentlyContinue
    throw 'SigmaNEST background worker is missing: '+$worker
  }
  $started=(Get-Date).ToUniversalTime()
  $initial=[ordered]@{
    jobId=$jobId
    state='STARTING'
    phase='QUEUED'
    message='SigmaNEST build accepted and queued.'
    workerVersion='2.11.1'
    pid=$null
    started=$started.ToString('o')
    finished=$null
    elapsedSeconds=0
    result=$null
    parts=@($request.reportParts)
    outputDir=[string]$request.outputDir
    jobName=[string]$request.jobName
    selectedSheets=@($request.selectedSheets)
  }
  Write-BuildStatus -statusFile $statusFile -obj $initial
  $psExe=Join-Path $PSHOME 'powershell.exe'
  if(-not(Test-Path -LiteralPath $psExe)){$psExe=(Get-Command powershell.exe -ErrorAction Stop).Source}
  $psi=New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName=$psExe
  $psi.Arguments='-NoProfile -ExecutionPolicy Bypass -STA -File "'+$worker+'" -RequestFile "'+$requestFile+'"'
  $psi.WorkingDirectory=$BridgeDir
  $psi.UseShellExecute=$false
  $psi.CreateNoWindow=$true
  $psi.RedirectStandardOutput=$false
  $psi.RedirectStandardError=$false
  try{
    $proc=[Diagnostics.Process]::Start($psi)
    $initial.pid=$proc.Id
    $initial.state='RUNNING'
    $initial.phase='STARTING'
    $initial.message='SigmaNEST background worker is running.'
    Write-BuildStatus -statusFile $statusFile -obj $initial
  }catch{
    Remove-Item -LiteralPath $requestFile -Force -ErrorAction SilentlyContinue
    throw ('Could not start SigmaNEST background worker: '+$_.Exception.Message)
  }
  [pscustomobject]@{jobId=$jobId;statusFile=$statusFile;requestFile=$requestFile;pid=$proc.Id;started=$started.ToString('o')}
}

function Read-BuildStatus([string]$jobId){
  $statusFile=Get-BuildStatusPath -jobId $jobId
  if(-not(Test-Path -LiteralPath $statusFile)){return $null}
  try{return (Get-Content -LiteralPath $statusFile -Raw -Encoding UTF8|ConvertFrom-Json)}
  catch{return [pscustomobject]@{jobId=$jobId;state='RUNNING';phase='STATUS_READ';message='Build status is being written; retry shortly.'}}
}
function Handle-Request($req){
  if($req.Method -eq 'OPTIONS'){return [pscustomobject]@{Status=204;Data=@{}}}
  if($req.Path -eq '/api/health' -and $req.Method -eq 'GET'){
    return [pscustomobject]@{Status=200;Data=@{ok=$true;port=$PORT;bridgeVersion=$BRIDGE_VERSION;libraryRoot=$CFG['libraryRoot'];dxfRoot=$CFG['dxfRoot'];lastScan=$CFG['lastScan'];count=[int]$CFG['count'];discoveredFiles=[int]$CFG['discoveredFiles'];prsCount=[int]$CFG['prsCount'];dxfCount=[int]$CFG['dxfCount'];dxfIndexState=[string]$CFG['dxfIndexState'];dxfIndexMessage=[string]$CFG['dxfIndexMessage'];sigmaNestCom=$true;bridge='PowerShell'}}
  }
  if($req.Path -eq '/api/scan' -and $req.Method -eq 'POST'){
    $b=if($req.Body){$req.Body|ConvertFrom-Json}else{[pscustomobject]@{}}
    $prsRoot=[IO.Path]::GetFullPath(([string]$(if($b.prsRoot){$b.prsRoot}else{$DEFAULT_LIBRARY})).Trim())
    $dxfRoot=[IO.Path]::GetFullPath(([string]$(if($b.dxfRoot){$b.dxfRoot}else{$DEFAULT_DXF_LIBRARY})).Trim())
    # Start the huge DXF walk in a background PowerShell process; never wait on Y:\ here.
    $dxfStatus=Start-DxfScan -root $dxfRoot
    $diag=Scan-Libraries -prsRoot $prsRoot -dxfRoot $dxfRoot
    return [pscustomobject]@{Status=200;Data=@{count=$diag.count;discoveredFiles=$diag.discoveredFiles;prsCount=$diag.prsCount;dxfCount=$diag.dxfCount;dxfIndexedCount=$dxfStatus.filesFound;dxfStatus=$dxfStatus;scanErrors=$diag.scanErrors;inspectErrors=$diag.inspectErrors;prsRoot=$diag.prsRoot;dxfRoot=$diag.dxfRoot}}
  }
  if($req.Path -eq '/api/dxf-status' -and $req.Method -eq 'GET'){
    $status=Get-DxfStatus
    return [pscustomobject]@{Status=200;Data=@{ok=$true;state=[string]$status.state;indexerVersion=[string]$status.indexerVersion;root=[string]$status.root;filesFound=[int]$status.filesFound;errors=[int]$status.errors;message=[string]$status.message;mode=[string]$status.mode;started=$status.started;finished=$status.finished;generatedUtc=$status.generatedUtc;currentPath=[string]$status.currentPath;pid=$(try{[int]$status.pid}catch{0});exitCode=$(try{[int]$status.exitCode}catch{0});workers=$(try{[int]$status.workers}catch{[int]$CFG['dxfWorkers']});workersCompleted=$(try{[int]$status.workersCompleted}catch{0});directoriesVisited=$(try{[int]$status.directoriesVisited}catch{0});elapsedSeconds=$(try{[double]$status.elapsedSeconds}catch{0});logFile=[string]$status.logFile;workDirectory=[string]$status.workDirectory;nextRefreshLocal=(Get-NextDxfRefreshLocal)}}
  }
  if($req.Path -eq '/api/dxf-refresh' -and $req.Method -eq 'POST'){
    $b=if($req.Body){$req.Body|ConvertFrom-Json}else{[pscustomobject]@{}}
    $dxfRoot=[IO.Path]::GetFullPath(([string]$(if($b.dxfRoot){$b.dxfRoot}else{$DEFAULT_DXF_LIBRARY})).Trim())
    $status=Get-DxfStatus
    if([string]$status.state -eq 'RUNNING'){
      return [pscustomobject]@{Status=200;Data=@{ok=$true;started=$false;dxfStatus=$status;message='A DXF indexing run is already active.'}}
    }
    $status=Start-DxfScan -root $dxfRoot -mode 'REFRESH' -force
    return [pscustomobject]@{Status=200;Data=@{ok=$true;started=$true;dxfStatus=$status;message='DXF refresh started in the background using controlled parallel workers.'}}
  }
  if($req.Path -eq '/api/build-job' -and $req.Method -eq 'POST'){
    $b=$req.Body|ConvertFrom-Json
    $prsRoot=[IO.Path]::GetFullPath(([string]$(if($b.prsRoot){$b.prsRoot}else{$DEFAULT_LIBRARY})).Trim())
    $dxfRoot=[IO.Path]::GetFullPath(([string]$(if($b.dxfRoot){$b.dxfRoot}else{$DEFAULT_DXF_LIBRARY})).Trim())
    $dxfStatus=Get-DxfStatus
    if([string]$dxfStatus.root -ne [string]$dxfRoot -or [string]$dxfStatus.state -ne 'COMPLETE'){
      $dxfStatus=Start-DxfScan -root $dxfRoot
    }
    # The server enforces the same rule as the UI: never classify unresolved
    # DXFs as missing/review while the massive Y:\ index is incomplete.
    if([string]$dxfStatus.root -eq [string]$dxfRoot -and [string]$dxfStatus.state -ne 'COMPLETE'){
      $statusCode=409
      return [pscustomobject]@{Status=$statusCode;Data=@{
        error='DXF index is not complete yet.'
        code='DXF_INDEX_NOT_READY'
        state=[string]$dxfStatus.state
        filesFound=[int]$dxfStatus.filesFound
        root=[string]$dxfStatus.root
        message=[string]$dxfStatus.message
        currentPath=[string]$dxfStatus.currentPath
      }}
    }
    if(
      -not $script:INDEX -or
      $script:INDEX.Count -eq 0 -or
      [string]$CFG['libraryRoot'] -ne [string]$prsRoot -or
      [string]$CFG['dxfRoot'] -ne [string]$dxfRoot -or
      ([string]$dxfStatus.state -eq 'COMPLETE' -and [string]$CFG['dxfIndexState'] -ne 'COMPLETE')
    ){
      Scan-Libraries -prsRoot $prsRoot -dxfRoot $dxfRoot|Out-Null
    }
    $jobRoot=$prsRoot
    $name=([string]$(if($b.jobName){$b.jobName}else{'CL_JOB'}) -replace '[^A-Za-z0-9._ -]','_').Trim();if(-not$name){$name='CL_JOB'}
    $parts=@()
    foreach($p in @($b.parts)){
      $f=Find-Part -part ([string]$p.part) -clMat ([string]$p.material) -clThk ([string]$p.thickness)
      if(-not$f){
        $ds=Get-DxfStatus
        if([string]$ds.root -eq [string]$dxfRoot -and [string]$ds.state -eq 'RUNNING'){
          Set-Prop $p 'status' 'REVIEW' | Out-Null
          Set-Prop $p 'statusLabel' 'DXF INDEXING' | Out-Null
          Set-Prop $p 'reviewReason' 'DXF server index is still being built; rerun the build when indexing completes' | Out-Null
        }else{
          Set-Prop $p 'status' 'MISSING' | Out-Null
          Set-Prop $p 'statusLabel' 'GEOMETRY MISSING' | Out-Null
          Set-Prop $p 'reviewReason' 'No matching .PRS or .DXF was found in the geometry library' | Out-Null
        }
        $parts+=$p
        continue
      }
      # Ambiguity is now metadata on a real selected geometry candidate.
      # Do not discard the source: reviewed parts must still be importable.
      $selectionAmbiguous=[bool]$(if($f.selectionAmbiguous){$f.selectionAmbiguous}else{$false})
      $selectionFuzzy=[bool]$(if($f.selectionFuzzy){$f.selectionFuzzy}else{$false})

      $sourceType=[string]$f.fileType
      $libMat=[string]$f.likelyMaterial
      $libThk=[string]$f.thickness
      $clMat=[string]$p.material
      $clThk=[string]$p.thickness
      $matKnown=!!$libMat.Trim()
      $thkKnown=!!$libThk.Trim()
      $clMatKnown=!!$clMat.Trim()
      $clThkKnown=!!$clThk.Trim()
      $mok=$true
      $tok=$true
      if($matKnown -and $clMatKnown){$mok=Material-Equal -a $clMat -b $libMat -aThickness $clThk -bThickness $libThk}
      if($thkKnown -and $clThkKnown){$tok=(Thickness-Number -s $clThk) -eq (Thickness-Number -s $libThk)}
      $variation=([string]$f.matchType -eq 'VARIATION')

      $reviewReason=''
      if($selectionAmbiguous){
        $status='REVIEW'
        $label='AMBIGUOUS ('+[int]$f.ambiguousCount+')'
        $scores=@($f.candidateScores|Select-Object -First 3)
        $reviewReason='Multiple geometry candidates matched; best candidate selected for import. '+(($scores|ForEach-Object{[IO.Path]::GetFileName([string]$_.file)+' score '+$_.score}) -join ' | ')
      }elseif($selectionFuzzy){
        $status='REVIEW'
        $label='FUZZY MATCH - REVIEW'
        $reviewReason='Geometry selected by conservative assembly-prefix matching; confirm the drawing before production.'
      }elseif([bool]$p.clReview){
        $status='REVIEW'
        $label=[string]$p.clReviewReason
        $reviewReason=[string]$(if($p.clReviewDetail){$p.clReviewDetail}else{'Conflicting CL values were kept separate and require confirmation'})
      }elseif(-not $clMatKnown -or -not $clThkKnown){
        $status='REVIEW'
        $label='CL MATERIAL/THICKNESS MISSING'
        $reviewReason='CL material or thickness is missing'
      }elseif($sourceType -eq 'DXF' -and $variation){
        $status='REVIEW'
        $label='VARIATION - DXF REVIEW'
        $reviewReason='The geometry was found as a DXF by a variation match and needs confirmation'
      }elseif($sourceType -eq 'PRS' -and $variation -and $matKnown -and $thkKnown -and $mok -and $tok){
        $status='READY'
        $label='FOUND - PRS VARIATION'
      }elseif($variation){
        $status='REVIEW'
        $label='VARIATION - REVIEW'
        $reviewReason='Variation match needs confirmation'
      }elseif($sourceType -eq 'DXF'){
        $status='READY'
        $label='FOUND - DXF'
      }elseif($matKnown -and -not $mok){
        $status='REVIEW'
        $label='MATERIAL MISMATCH'
        $reviewReason='PRS material conflicts with CL material'
      }elseif($thkKnown -and -not $tok){
        $status='REVIEW'
        $label='THICKNESS MISMATCH'
        $reviewReason='PRS thickness conflicts with CL thickness'
      }else{
        $status='READY'
        $label=if($matKnown -and $thkKnown){'FOUND - PRS'}else{'FOUND - PRS USING CL DATA'}
      }

      Set-Prop -obj $p -name 'status' -value $status | Out-Null
      Set-Prop -obj $p -name 'statusLabel' -value $label | Out-Null
      Set-Prop -obj $p -name 'reviewReason' -value $reviewReason | Out-Null
      Set-Prop -obj $p -name 'sourcePath' -value ([string]$f.file) | Out-Null
      Set-Prop -obj $p -name 'sourceType' -value $sourceType | Out-Null
      Set-Prop -obj $p -name 'file' -value ([string]$f.file) | Out-Null
      Set-Prop -obj $p -name 'prs' -value $(if($sourceType -eq 'PRS'){[string]$f.file}else{''}) | Out-Null
      Set-Prop -obj $p -name 'sourceDxf' -value $(if($sourceType -eq 'DXF'){[string]$f.file}else{[string]$f.sourceDxf}) | Out-Null
      Set-Prop -obj $p -name 'matchType' -value $f.matchType | Out-Null
      Set-Prop $p 'libraryMaterial' $libMat | Out-Null
      Set-Prop $p 'libraryThickness' $libThk | Out-Null
      $parts+=$p
    }
    $review=@($parts|Where-Object {$_.status -ne 'READY'})
    $reviewBreakdown=[ordered]@{}
    foreach($rp in $review){
      $reason=[string]$rp.statusLabel
      if([string]::IsNullOrWhiteSpace($reason)){$reason='REVIEW'}
      if(-not $reviewBreakdown.Contains($reason)){$reviewBreakdown[$reason]=0}
      $reviewBreakdown[$reason]=[int]$reviewBreakdown[$reason]+1
    }
    $staging=Write-Job -root $jobRoot -name $name -parts $parts

    # Review no longer blocks WS creation. Any part with a resolved geometry
    # source is handed to SigmaNEST, even when it is marked REVIEW so the
    # operator can correct it directly in SigmaNEST. Only true MISSING rows
    # have no geometry to import.
    $importable=@($parts|Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_.sourcePath)})
    $missing=@($parts|Where-Object {$_.status -eq 'MISSING'})
    if($importable.Count -eq 0){
      return [pscustomobject]@{
        Status=200
        Data=@{
          outputDir=$staging
          message="No geometry could be imported into SigmaNEST. $($missing.Count) part(s) have missing geometry."
          parts=$parts
          reviewCount=$review.Count
          reviewBreakdown=$reviewBreakdown
          sigmaNestCreated=$false
          sigmaPartCount=0
          importedCount=0
          missingCount=$missing.Count
          taskPlan=@()
        }
      }
    }

    $reportParts=@($parts)
    $taskPlan=@()
    foreach($p0 in @($parts)){
      if($p0.taskBatches){
        foreach($tb in @($p0.taskBatches)){
          $taskPlan += [pscustomobject]@{
            sheet=[string]$tb.sheet
            batchMultiplier=$(if($tb.batchMultiplier){$tb.batchMultiplier}else{1})
            part=[string]$p0.part
            qty=[double]$(if($tb.vehicleQty){$tb.vehicleQty}else{$p0.qty})
          }
        }
      }else{
        $taskPlan += [pscustomobject]@{
          sheet=[string]$p0.sheet
          batchMultiplier=$(if($p0.batchMultiplier){$p0.batchMultiplier}else{1})
          part=[string]$p0.part
          qty=[double]$(if($p0.vehicleQty){$p0.vehicleQty}else{$p0.qty})
        }
      }
    }
    $engineRequest=[pscustomobject]@{
      jobName=$name
      libraryRoot=$jobRoot
      dxfRoot=$dxfRoot
      wsDirectory=[string]$b.wsDirectory
      parts=@($importable|ForEach-Object{
        [pscustomobject]@{
          part=$_.part
          qty=$_.qty
          batchMultiplier=$(if($_.batchMultiplier){$_.batchMultiplier}else{1})
          taskBatches=@($_.taskBatches)
          taskSheet=$_.sheet
          sourcePath=$_.sourcePath
          sourceType=$_.sourceType
          prsPath=$(if($_.sourceType -eq 'PRS'){[string]$_.sourcePath}else{''})
          sigmaMaterial=(Sigma-Material -cl ([string]$_.material) -lib ([string]$_.libraryMaterial))
          thicknessMm=(Thickness-Number -s ([string]$_.thickness))
        }
      })
    }
    $workerRequest=[pscustomobject]@{
      jobId=''
      statusFile=''
      jobName=$name
      outputDir=$staging
      selectedSheets=@($b.selectedSheetNames)
      reportParts=$reportParts
      engineRequest=$engineRequest
    }
    $workerInfo=Start-SigmaNestBuildWorker -request $workerRequest
    $reviewText=if($review.Count -gt 0){
      "SigmaNEST job accepted. $($importable.Count) of $($parts.Count) CL part(s) have geometry and are being imported in the background. $($review.Count) part(s) require attention and will be listed in Part Review."
    }else{
      "SigmaNEST job accepted. All $($parts.Count) CL part(s) have geometry and are being imported in the background."
    }
    return [pscustomobject]@{
      Status=202
      Data=@{
        accepted=$true
        state='RUNNING'
        jobId=$workerInfo.jobId
        pid=$workerInfo.pid
        started=$workerInfo.started
        outputDir=$staging
        message=$reviewText
        parts=$parts
        reviewCount=$review.Count
        reviewBreakdown=$reviewBreakdown
        sigmaNestCreated=$false
        wsPath=''
        sigmaPartCount=0
        importedCount=0
        missingCount=$missing.Count
        taskPlan=@($taskPlan)
      }
    }
  if($req.Path -match '^/api/build-status/([0-9a-fA-F-]{36})$' -and $req.Method -eq 'GET'){
    $jobId=$Matches[1]
    $status=Read-BuildStatus -jobId $jobId
    if($null -eq $status){return [pscustomobject]@{Status=404;Data=@{error='Build job not found.';jobId=$jobId}}}
    return [pscustomobject]@{Status=200;Data=$status}
  }
  [pscustomobject]@{Status=404;Data=@{error='Not found'}}
}

$listener = New-Object -TypeName Net.Sockets.TcpListener -ArgumentList ([Net.IPAddress]::Loopback,$PORT)
$listener.Start()
Write-Host "CL-WS-Creator PowerShell bridge listening on http://127.0.0.1:$PORT"
Write-Host "PRS library default: $DEFAULT_LIBRARY"
Write-Host "DXF library default: $DEFAULT_DXF_LIBRARY"
Write-Host "DXF workers: $(Get-DxfWorkerCount); nightly refresh hour: $($CFG['dxfNightlyHour'])"
Initialize-DxfScheduler
$lastSchedulerCheck=Get-Date
while($true){
  $acceptTask=$listener.AcceptTcpClientAsync()
  while(-not $acceptTask.Wait(1000)){
    if(((Get-Date)-$lastSchedulerCheck).TotalSeconds -ge 30){
      Invoke-DxfScheduler
      $lastSchedulerCheck=Get-Date
    }
  }
  if(((Get-Date)-$lastSchedulerCheck).TotalSeconds -ge 30){
    Invoke-DxfScheduler
    $lastSchedulerCheck=Get-Date
  }
  $client=$acceptTask.Result
  try{
    $req=Read-HttpRequest $client
    try{$resp=Handle-Request $req}catch{$resp=[pscustomobject]@{Status=500;Data=@{error=$_.Exception.Message}}}
    Send-Json $client $resp.Status $resp.Data
  }catch{
    try{Send-Json $client 500 @{error=$_.Exception.Message}}catch{}
  }
}
