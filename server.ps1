param()

$ErrorActionPreference = 'Stop'
$PORT = 17832
$BRIDGE_VERSION = '2.3.1'
$DEFAULT_LIBRARY = if($env:SN_PARTS){$env:SN_PARTS}else{'S:\SNDataX1\PARTS'}
$DEFAULT_DXF_LIBRARY = if($env:SN_DXF){$env:SN_DXF}else{'Y:\'}
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG_FILE = Join-Path $ROOT 'config.json'
$DXF_INDEX_FILE = Join-Path $ROOT 'dxf-index.json'
$DXF_INDEX_DIR = Join-Path $ROOT 'dxf-index'
$DXF_STATUS_FILE = Join-Path $ROOT 'dxf-index-status.json'
$DXF_SCAN_SCRIPT = Join-Path $ROOT 'dxf-indexer.ps1'
$CFG = [ordered]@{ libraryRoot=$DEFAULT_LIBRARY; dxfRoot=$DEFAULT_DXF_LIBRARY; lastScan=$null; count=0; discoveredFiles=0; prsCount=0; dxfCount=0; dxfIndexState='IDLE'; dxfIndexMessage=''; scanErrors=@(); inspectErrors=@() }

try {
  if(Test-Path -LiteralPath $CFG_FILE){
    $old = Get-Content -LiteralPath $CFG_FILE -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach($p in $old.PSObject.Properties){ $CFG[$p.Name] = $p.Value }
  }
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

function Stop-RunningDxfWorkers([string]$root){
  try{
    $rootPattern=[regex]::Escape($root)
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
      Where-Object {
        $_.ProcessId -ne $PID -and
        ([string]$_.CommandLine -match 'dxf-indexer\.ps1') -and
        ([string]$_.CommandLine -match $rootPattern)
      } |
      ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
      }
  }catch{}
}

function Start-DxfScan([string]$root){
  $status=Get-DxfStatus
  $requiredIndexerVersion='2.0.2'

  if(([string]$status.root -eq [string]$root) -and [string]$status.state -eq 'COMPLETE' -and
     [string]$status.indexerVersion -eq $requiredIndexerVersion -and
     (Test-Path -LiteralPath $DXF_INDEX_FILE) -and
     (Test-Path -LiteralPath $DXF_INDEX_DIR)){
    return $status
  }

  if(([string]$status.root -eq [string]$root) -and [string]$status.state -eq 'RUNNING' -and
     [string]$status.indexerVersion -eq $requiredIndexerVersion){
    try{
      if($status.pid){
        $p=Get-Process -Id ([int]$status.pid) -ErrorAction SilentlyContinue
        if($p){return $status}
      }
    }catch{}
  }

  # Only stop workers when the recorded worker is stale, missing, or belongs
  # to an older indexer/root. A healthy current worker must never be restarted.
  Stop-RunningDxfWorkers -root $root

  if(-not(Test-Path -LiteralPath $DXF_SCAN_SCRIPT)){
    return [pscustomobject]@{
      state='FAILED'
      indexerVersion=$requiredIndexerVersion
      root=$root
      filesFound=0
      errors=1
      message='DXF indexer script is missing: '+$DXF_SCAN_SCRIPT
    }
  }

  try{
    Remove-Item -LiteralPath ($DXF_INDEX_FILE+'.tmp') -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath ($DXF_INDEX_DIR+'.tmp') -Recurse -Force -ErrorAction SilentlyContinue

    $started=(Get-Date).ToUniversalTime().ToString('o')
    $starting=[pscustomobject]@{
      state='RUNNING'
      indexerVersion=$requiredIndexerVersion
      root=$root
      message='DXF indexing worker is starting.'
      filesFound=0
      errors=0
      started=$started
      finished=$null
      currentPath=$root
      pid=$null
    }
    ($starting|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8

    # Launch the worker through System.Diagnostics.Process so its
    # stdout/stderr can be captured even when Windows PowerShell exits before
    # the script gets far enough to create its own log.
    $command="& '$DXF_SCAN_SCRIPT' -Root '$root' -IndexFile '$DXF_INDEX_FILE' -StatusFile '$DXF_STATUS_FILE'"
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $launchLog=Join-Path $ROOT 'dxf-indexer-launch.log'
    $stdoutLog=Join-Path $ROOT 'dxf-indexer-stdout.log'
    $stderrLog=Join-Path $ROOT 'dxf-indexer-stderr.log'
    try{
      ('['+(Get-Date).ToUniversalTime().ToString('o')+'] Launching powershell.exe via ProcessStartInfo. Script='+$DXF_SCAN_SCRIPT+' Root='+$root)|Add-Content -LiteralPath $launchLog -Encoding UTF8
      $psi=New-Object System.Diagnostics.ProcessStartInfo
      $psi.FileName='powershell.exe'
      $psi.Arguments='-NoProfile -ExecutionPolicy Bypass -EncodedCommand '+$encoded
      $psi.WorkingDirectory=$ROOT
      $psi.UseShellExecute=$false
      $psi.CreateNoWindow=$true
      $psi.RedirectStandardOutput=$true
      $psi.RedirectStandardError=$true
      $proc=[System.Diagnostics.Process]::Start($psi)
    }catch{
      $message='Could not launch DXF indexer process: '+$_.Exception.Message
      $failed=[pscustomobject]@{
        state='FAILED'
        indexerVersion=$requiredIndexerVersion
        root=$root
        filesFound=0
        errors=1
        message=$message
        pid=$null
        logFile=$stderrLog
        launchLogFile=$launchLog
      }
      try{($failed|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8}catch{}
      return $failed
    }

    try{
      $starting.pid=$proc.Id
      $starting.message='DXF indexer process launched (PID '+$proc.Id+'); waiting for the indexer heartbeat.'
      $starting.logFile=$stdoutLog
      $starting.launchLogFile=$launchLog
      ($starting|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8
      ('['+(Get-Date).ToUniversalTime().ToString('o')+'] Process launched. PID='+$proc.Id)|Add-Content -LiteralPath $launchLog -Encoding UTF8
    }catch{}

    Start-Sleep -Milliseconds 2500
    try{
      if($proc.HasExited){
        $stdout=''
        $stderr=''
        try{$stdout=$proc.StandardOutput.ReadToEnd()}catch{}
        try{$stderr=$proc.StandardError.ReadToEnd()}catch{}
        try{if($stdout){$stdout|Set-Content -LiteralPath $stdoutLog -Encoding UTF8}}catch{}
        try{if($stderr){$stderr|Set-Content -LiteralPath $stderrLog -Encoding UTF8}}catch{}
        $exitCode=$proc.ExitCode
        $parts=@('DXF indexer exited immediately after launch (exit code '+$exitCode+').')
        if($stderr){$parts+=('STDERR: '+(($stderr -replace '\r?\n',' ') -replace '\s+',' ').Trim())}
        if($stdout){$parts+=('STDOUT: '+(($stdout -replace '\r?\n',' ') -replace '\s+',' ').Trim())}
        $message=$parts -join ' '
        $failed=[pscustomobject]@{
          state='FAILED'
          indexerVersion=$requiredIndexerVersion
          root=$root
          filesFound=0
          errors=1
          message=$message
          pid=$proc.Id
          exitCode=$exitCode
          logFile=$stdoutLog
          stderrLogFile=$stderrLog
          launchLogFile=$launchLog
        }
        try{($failed|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8}catch{}
        return $failed
      }

      try{
        $current=Get-Content -LiteralPath $DXF_STATUS_FILE -Raw -Encoding UTF8|ConvertFrom-Json
        $current.message='DXF indexer process is alive (PID '+$proc.Id+'). Waiting for the first scan heartbeat.'
        $current.logFile=$stdoutLog
        $current.stderrLogFile=$stderrLog
        $current.launchLogFile=$launchLog
        ($current|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8
        ('['+(Get-Date).ToUniversalTime().ToString('o')+'] Process still alive after 2.5 seconds. PID='+$proc.Id)|Add-Content -LiteralPath $launchLog -Encoding UTF8
      }catch{}
    }catch{}

    return $starting
  }catch{
    $failed=[pscustomobject]@{
      state='FAILED'
      indexerVersion=$requiredIndexerVersion
      root=$root
      filesFound=0
      errors=1
      message=$_.Exception.Message
    }
    try{($failed|ConvertTo-Json -Depth 8)|Set-Content -LiteralPath $DXF_STATUS_FILE -Encoding UTF8}catch{}
    return $failed
  }
}
function Dxf-ShardKey([string]$part){
  $n=Normalize -s $part
  if([string]::IsNullOrWhiteSpace($n)){return '__'}
  if($n.Length -ge 2){return $n.Substring(0,2)}
  return $n+'_'
}

function Get-DxfShardCandidates([string]$root,[string]$part){
  if([string]::IsNullOrWhiteSpace($root)){return @()}
  $status=Get-DxfStatus
  if([string]$status.state -ne 'COMPLETE' -or [string]$status.root -ne [string]$root -or [string]$status.indexerVersion -ne '2.0.2'){return @()}
  if(-not(Test-Path -LiteralPath $DXF_INDEX_DIR)){return @()}

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
  $prsFiles=@(Get-GeometryFiles -root $prsRoot -extension '.prs' -errors $scanErrors)
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
  $CFG.libraryRoot=$prsRoot
  $CFG.dxfRoot=$dxfRoot
  $CFG.lastScan=(Get-Date).ToUniversalTime().ToString('o')
  $CFG.count=$total
  $CFG.discoveredFiles=$total
  $CFG.prsCount=$prsFiles.Count
  $CFG.dxfCount=$dxfCount
  $CFG.dxfIndexState=[string]$dxfStatus.state
  $CFG.dxfIndexMessage=[string]$dxfStatus.message
  $CFG.scanErrors=@($scanErrors)
  $CFG.inspectErrors=@()
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

function Select-MatchCandidate($candidates) {
  $c=@($candidates)
  if($c.Count -eq 1){return $c[0]}
  if($c.Count -gt 1){
    # Prefer one unique PRS when the same part is also present as a DXF.
    $prs=@($c|Where-Object {$_.fileType -eq 'PRS'})
    if($prs.Count -eq 1){return $prs[0]}
    return [pscustomobject]@{ambiguous=$c}
  }
  $null
}

function Find-Part([string]$part){
  $n=Normalize -s $part
  $dxf=@()
  if($script:DXF_ROOT){$dxf=@(Get-DxfShardCandidates -root $script:DXF_ROOT -part $part)}

  # Match in stages while combining PRS and DXF candidates at each stage.
  # A unique PRS candidate is preferred when it is present alongside a DXF
  # candidate for the same exact match.
  $prsExact=@()
  if($script:BYNAME -and $script:BYNAME.ContainsKey($n)){$prsExact=@($script:BYNAME[$n])}
  $dxfExact=@($dxf|Where-Object {(Normalize -s $_.partName) -eq $n})
  $exact=@($prsExact+$dxfExact)
  if($exact.Count){
    $hit=Select-MatchCandidate $exact
    if($hit -and -not($hit.PSObject.Properties.Name -contains 'ambiguous')){
      Ensure-Metadata -item $hit | Out-Null
      Set-Prop $hit 'matchType' 'EXACT' | Out-Null
      return $hit
    }
    if($hit){return $hit}
  }

  $prsEmbedded=@($script:INDEX|Where-Object {Is-EmbeddedMatch -candidate $_.embeddedPartName -target $part})
  $dxfEmbedded=@($dxf|Where-Object {Is-EmbeddedMatch -candidate $_.embeddedPartName -target $part})
  $embedded=@($prsEmbedded+$dxfEmbedded)
  if($embedded.Count){
    $hit=Select-MatchCandidate $embedded
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
  $vars=@($prsVars+$dxfVars)
  if($vars.Count){
    $hit=Select-MatchCandidate $vars
    if($hit -and -not($hit.PSObject.Properties.Name -contains 'ambiguous')){
      Ensure-Metadata -item $hit | Out-Null
      Set-Prop $hit 'matchType' 'VARIATION' | Out-Null
      return $hit
    }
    if($hit){return $hit}
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
    $dest=Join-Path -Path $partsDir -ChildPath ([IO.Path]::GetFileName([string]$p.sourcePath))
    Copy-Item -LiteralPath ([string]$p.sourcePath) -Destination $dest -Force
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

function Handle-Request($req){
  if($req.Method -eq 'OPTIONS'){return [pscustomobject]@{Status=204;Data=@{}}}
  if($req.Path -eq '/api/health' -and $req.Method -eq 'GET'){
    return [pscustomobject]@{Status=200;Data=@{ok=$true;port=$PORT;bridgeVersion=$BRIDGE_VERSION;libraryRoot=$CFG.libraryRoot;dxfRoot=$CFG.dxfRoot;lastScan=$CFG.lastScan;count=[int]$CFG.count;discoveredFiles=[int]$CFG.discoveredFiles;prsCount=[int]$CFG.prsCount;dxfCount=[int]$CFG.dxfCount;dxfIndexState=[string]$CFG.dxfIndexState;dxfIndexMessage=[string]$CFG.dxfIndexMessage;sigmaNestCom=$true;bridge='PowerShell'}}
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
    return [pscustomobject]@{Status=200;Data=@{ok=$true;state=[string]$status.state;indexerVersion=[string]$status.indexerVersion;root=[string]$status.root;filesFound=[int]$status.filesFound;errors=[int]$status.errors;message=[string]$status.message;started=$status.started;finished=$status.finished;currentPath=[string]$status.currentPath;pid=$(try{[int]$status.pid}catch{0});logFile=[string]$status.logFile}}
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
      [string]$CFG.libraryRoot -ne [string]$prsRoot -or
      [string]$CFG.dxfRoot -ne [string]$dxfRoot -or
      ([string]$dxfStatus.state -eq 'COMPLETE' -and [string]$CFG.dxfIndexState -ne 'COMPLETE')
    ){
      Scan-Libraries -prsRoot $prsRoot -dxfRoot $dxfRoot|Out-Null
    }
    $root=$prsRoot
    $name=([string]$(if($b.jobName){$b.jobName}else{'CL_JOB'}) -replace '[^A-Za-z0-9._ -]','_').Trim();if(-not$name){$name='CL_JOB'}
    $parts=@()
    foreach($p in @($b.parts)){
      $f=Find-Part -part ([string]$p.part)
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
      if($f.PSObject.Properties.Name -contains 'ambiguous'){
        Set-Prop $p 'status' 'REVIEW' | Out-Null
        Set-Prop $p 'statusLabel' ('AMBIGUOUS ('+$f.ambiguous.Count+')') | Out-Null
        Set-Prop $p 'reviewReason' 'More than one geometry source matches this part' | Out-Null
        $parts+=$p
        continue
      }

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
      if([bool]$p.clReview){
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
    $staging=Write-Job -root $root -name $name -parts $parts
    if($review.Count -gt 0){return [pscustomobject]@{Status=200;Data=@{outputDir=$staging;message="Job staged, but $($review.Count) part(s) require review before SigmaNEST creation.";parts=$parts;reviewCount=$review.Count;reviewBreakdown=$reviewBreakdown;sigmaNestCreated=$false}}}
    $reqFile=Join-Path $ROOT ('_psrequest-'+[Diagnostics.Process]::GetCurrentProcess().Id+'-'+[DateTime]::Now.Ticks+'.json')
    $request=[pscustomobject]@{jobName=$name;libraryRoot=$root;dxfRoot=$dxfRoot;wsDirectory=[string]$b.wsDirectory;parts=@($parts|ForEach-Object{[pscustomobject]@{part=$_.part;qty=$_.qty;batchMultiplier=$(if($_.batchMultiplier){$_.batchMultiplier}else{1});taskSheet=$_.sheet;sourcePath=$_.sourcePath;sourceType=$_.sourceType;prsPath=$(if($_.sourceType -eq 'PRS'){[string]$_.sourcePath}else{''});sigmaMaterial=(Sigma-Material -cl ([string]$_.material) -lib ([string]$_.libraryMaterial));thicknessMm=(Thickness-Number -s ([string]$_.thickness))}})}
    ($request|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath $reqFile -Encoding UTF8
    try{
      $worker=Join-Path $ROOT 'create-sigmanest-ws.ps1'
      $raw=& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $worker -RequestFile $reqFile 2>&1 | Out-String
      $data=$raw.Trim()|ConvertFrom-Json
      if(-not$data.ok){throw $data.error}
      return [pscustomobject]@{Status=200;Data=@{outputDir=$staging;message=$data.message;parts=$parts;reviewCount=0;sigmaNestCreated=$true;wsPath=$data.wsPath;sigmaPartCount=$data.partCount;taskPlan=@($parts|ForEach-Object{[pscustomobject]@{sheet=$_.sheet;batchMultiplier=$(if($_.batchMultiplier){$_.batchMultiplier}else{1});part=$_.part;qty=$_.qty}})}}
    }finally{Remove-Item -LiteralPath $reqFile -Force -ErrorAction SilentlyContinue}
  }
  [pscustomobject]@{Status=404;Data=@{error='Not found'}}
}

$listener = New-Object -TypeName Net.Sockets.TcpListener -ArgumentList ([Net.IPAddress]::Loopback,$PORT)
$listener.Start()
Write-Host "CL-WS-Creator PowerShell bridge listening on http://127.0.0.1:$PORT"
Write-Host "PRS library default: $DEFAULT_LIBRARY"
while($true){
  $client=$listener.AcceptTcpClient()
  try{
    $req=Read-HttpRequest $client
    try{$resp=Handle-Request $req}catch{$resp=[pscustomobject]@{Status=500;Data=@{error=$_.Exception.Message}}}
    Send-Json $client $resp.Status $resp.Data
  }catch{
    try{Send-Json $client 500 @{error=$_.Exception.Message}}catch{}
  }
}
