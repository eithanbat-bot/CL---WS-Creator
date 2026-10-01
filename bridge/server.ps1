param()

$ErrorActionPreference = 'Stop'
$PORT = 17832
$DEFAULT_LIBRARY = if($env:SN_PARTS){$env:SN_PARTS}else{'S:\SNDataX1\PARTS'}
$ROOT = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG_FILE = Join-Path $ROOT 'config.json'
$CFG = [ordered]@{ libraryRoot=$DEFAULT_LIBRARY; lastScan=$null; count=0; discoveredFiles=0; scanErrors=@(); inspectErrors=@() }

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
function VariationKey([string]$s){ $n=Normalize $s; if($n -match '[A-Z]$'){$n=$n.Substring(0,$n.Length-1)}; $n }
function Normalize-Material([string]$s){
  if($null -eq $s){$s=''}
  $s=$s -replace '(?i)Armoxt|Amoxt','Armox'
  $s=$s -replace '(?i)Ramort','Ramor'
  ($s -replace '\s+',' ').Trim()
}
function Thickness-Number([string]$s){
  $m=[regex]::Match([string]$s,'(\d+(?:\.\d+)?)\s*mm','IgnoreCase')
  if($m.Success){[double]$m.Groups[1].Value}else{[double]::NaN}
}
function Material-Equal([string]$a,[string]$b){
  $A=(Normalize-Material $a).ToUpperInvariant();$B=(Normalize-Material $b).ToUpperInvariant()
  if(-not $A -or -not $B -or $A -eq $B){return $true}
  if($A -match '4\s*MM\s*ARMOX' -and $B -match 'RAMOR\s*500'){return $true}
  if($B -match '4\s*MM\s*ARMOX' -and $A -match 'RAMOR\s*500'){return $true}
  return ($A.Contains($B) -or $B.Contains($A))
}
function Sigma-Material([string]$cl,[string]$lib){
  $s=Normalize-Material $(if($lib){$lib}else{$cl})
  $s=$s -replace '^\d+(?:\.\d+)?\s*mm\s*',''
  $s=$s -replace '\s+(sheet|plate)$',''
  $s.Trim()
}
function Scan-Library([string]$root) {
  if(-not(Test-Path -LiteralPath $root)){throw "Cannot access the PRS library folder: $root"}
  if(-not((Get-Item -LiteralPath $root).PSIsContainer)){throw "Library path is not a folder: $root"}
  $files=@();$scanErrors=@();$inspectErrors=@()
  try{$files=@(Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.Extension -ieq '.prs'} | ForEach-Object {$_.FullName})}
  catch{ $scanErrors += [pscustomobject]@{path=$root;error=$_.Exception.Message} }
  $items=@()
  foreach($file in $files){
    try{$items += Inspect-Prs $file}catch{$inspectErrors += [pscustomobject]@{path=$file;error=$_.Exception.Message}}
  }
  $script:INDEX=$items
  $CFG.libraryRoot=$root; $CFG.lastScan=(Get-Date).ToUniversalTime().ToString('o');$CFG.count=$items.Count
  $CFG.discoveredFiles=$files.Count;$CFG.scanErrors=$scanErrors;$CFG.inspectErrors=$inspectErrors
  try{($CFG|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath $CFG_FILE -Encoding UTF8}catch{}
  [pscustomobject]@{count=$items.Count;discoveredFiles=$files.Count;scanErrors=$scanErrors;inspectErrors=$inspectErrors;root=$root}
}
function Find-Part([string]$part){
  $n=Normalize $part
  $h=$script:INDEX|Where-Object {(Normalize $_.partName) -eq $n}|Select-Object -First 1
  if($h){$h|Add-Member NoteProperty matchType EXACT -Force;return $h}
  $h=$script:INDEX|Where-Object {(Normalize $_.embeddedPartName) -eq $n -or (Normalize $_.embeddedPartName).StartsWith($n+'-')}|Select-Object -First 1
  if($h){$h|Add-Member NoteProperty matchType EMBEDDED -Force;return $h}
  $vk=VariationKey $part;$vars=@($script:INDEX|Where-Object {(VariationKey $_.partName)-eq $vk})
  if($vars.Count -eq 1){$vars[0]|Add-Member NoteProperty matchType VARIATION -Force;return $vars[0]}
  if($vars.Count -gt 1){return [pscustomobject]@{ambiguous=$vars}}
  $null
}
function Csv([object]$v){
  $s=[string]$v
  if($s -match '[,"\r\n]'){'"' + ($s -replace '"','""') + '"'}else{$s}
}
function Write-Job([string]$root,[string]$name,$parts){
  $out=Join-Path $root ('_CL_WS_BUILDER\'+$name)
  New-Item -ItemType Directory -Path (Join-Path $out 'parts') -Force|Out-Null
  foreach($p in @($parts|Where-Object {$_.file})){Copy-Item -LiteralPath $p.file -Destination (Join-Path $out 'parts' ([IO.Path]::GetFileName($p.file))) -Force}
  $rows=@('Part,Qty,Material,Thickness,PRS,MatchType,PRS_Material,PRS_Thickness,SourceDXF,Status')
  foreach($p in $parts){
    $vals=@($p.part,$p.qty,$p.material,$p.thickness,$(if($p.prs){[IO.Path]::GetFileName($p.prs)}else{''}),$p.matchType,$p.libraryMaterial,$p.libraryThickness,$p.sourceDxf,$p.status)
    $rows += (($vals|ForEach-Object{Csv $_}) -join ',')
  }
  $rows|Set-Content -LiteralPath (Join-Path $out 'WS_PARTS.csv') -Encoding UTF8
  ($parts|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath (Join-Path $out 'SIGMANEST_JOB.json') -Encoding UTF8
  $out
}

function Handle-Request($req){
  if($req.Method -eq 'OPTIONS'){return [pscustomobject]@{Status=204;Data=@{}}}
  if($req.Path -eq '/api/health' -and $req.Method -eq 'GET'){
    return [pscustomobject]@{Status=200;Data=@{ok=$true;port=$PORT;libraryRoot=$CFG.libraryRoot;lastScan=$CFG.lastScan;count=[int]$CFG.count;discoveredFiles=[int]$CFG.discoveredFiles;sigmaNestCom=$true;bridge='PowerShell'}}
  }
  if($req.Path -eq '/api/scan' -and $req.Method -eq 'POST'){
    $b=if($req.Body){$req.Body|ConvertFrom-Json}else{[pscustomobject]@{}}
    $root=[IO.Path]::GetFullPath(([string]$(if($b.root){$b.root}else{$DEFAULT_LIBRARY})).Trim())
    $diag=Scan-Library $root
    return [pscustomobject]@{Status=200;Data=@{count=$diag.count;discoveredFiles=$diag.discoveredFiles;scanErrors=$diag.scanErrors;inspectErrors=$diag.inspectErrors;root=$diag.root;parts=@($script:INDEX|ForEach-Object{[pscustomobject]@{partName=$_.partName;embeddedPartName=$_.embeddedPartName;likelyMaterial=$_.likelyMaterial;thickness=$_.thickness;sourceDxf=$_.sourceDxf}})}}
  }
  if($req.Path -eq '/api/build-job' -and $req.Method -eq 'POST'){
    $b=$req.Body|ConvertFrom-Json
    $root=[IO.Path]::GetFullPath(([string]$(if($b.libraryRoot){$b.libraryRoot}else{$DEFAULT_LIBRARY})).Trim())
    Scan-Library $root|Out-Null
    $name=([string]$(if($b.jobName){$b.jobName}else{'CL_JOB'}) -replace '[^A-Za-z0-9._ -]','_').Trim();if(-not$name){$name='CL_JOB'}
    $parts=@()
    foreach($p in @($b.parts)){
      $f=Find-Part ([string]$p.part)
      if(-not$f){$p|Add-Member NoteProperty status MISSING -Force;$p|Add-Member NoteProperty statusLabel 'GEOMETRY MISSING' -Force;$parts+=$p;continue}
      if($f.PSObject.Properties.Name -contains 'ambiguous'){
        $p|Add-Member NoteProperty status REVIEW -Force;$p|Add-Member NoteProperty statusLabel ('AMBIGUOUS ('+$f.ambiguous.Count+')') -Force;$parts+=$p;continue
      }
      $libMat=[string]$f.likelyMaterial;$libThk=[string]$f.thickness
      $matKnown=!!$libMat.Trim();$thkKnown=!!$libThk.Trim();$mok=$matKnown -and (Material-Equal $p.material $libMat);$tok=$thkKnown -and !!$p.thickness -and ((Thickness-Number $p.thickness) -eq (Thickness-Number $libThk));$variation=([string]$f.matchType -eq 'VARIATION')
      $status=if($matKnown -and $thkKnown -and $mok -and $tok -and -not$variation){'READY'}else{'REVIEW'}
      $label=if($status -eq 'READY'){'FOUND'}elseif($variation){'VARIATION - REVIEW'}elseif(-not$matKnown){'MATERIAL NOT CONFIRMED'}elseif(-not$mok){'MATERIAL MISMATCH'}elseif(-not$thkKnown){'THICKNESS NOT CONFIRMED'}elseif(-not$tok){'THICKNESS MISMATCH'}else{'REVIEW'}
      $p|Add-Member NoteProperty status $status -Force;$p|Add-Member NoteProperty statusLabel $label -Force
      foreach($kv in @{'file'=$f.file;'prs'=$f.file;'sourceDxf'=$f.sourceDxf;'matchType'=$f.matchType;'libraryMaterial'=$libMat;'libraryThickness'=$libThk}){$p|Add-Member NoteProperty $kv.Key $kv.Value -Force}
      $parts+=$p
    }
    $review=@($parts|Where-Object {$_.status -ne 'READY'})
    $staging=Write-Job $root $name $parts
    if($review.Count -gt 0){return [pscustomobject]@{Status=200;Data=@{outputDir=$staging;message="Job staged, but $($review.Count) part(s) require review before SigmaNEST creation.";parts=$parts;reviewCount=$review.Count;sigmaNestCreated=$false}}}
    $reqFile=Join-Path $ROOT ('_psrequest-'+[Diagnostics.Process]::GetCurrentProcess().Id+'-'+[DateTime]::Now.Ticks+'.json')
    $request=[pscustomobject]@{jobName=$name;libraryRoot=$root;wsDirectory=[string]$b.wsDirectory;parts=@($parts|ForEach-Object{[pscustomobject]@{part=$_.part;qty=$_.qty;batchMultiplier=$(if($_.batchMultiplier){$_.batchMultiplier}else{1});taskSheet=$_.sheet;prsPath=$_.file;sigmaMaterial=(Sigma-Material $_.material $_.libraryMaterial);thicknessMm=(Thickness-Number $_.thickness)}})}
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

$listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,$PORT)
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
