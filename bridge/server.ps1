param()

$ErrorActionPreference = 'Stop'
$PORT = 17832
$BRIDGE_VERSION = '2.1.0'
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
function VariationKey([string]$s){ $n=Normalize -s $s; if($n -match '[A-Z]$'){$n=$n.Substring(0,$n.Length-1)}; $n }
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
function Material-Equal([string]$a,[string]$b){
  $A=(Normalize-Material -s $a).ToUpperInvariant()
  $B=(Normalize-Material -s $b).ToUpperInvariant()
  if(-not $A -or -not $B -or $A -eq $B){return $true}

  # Required job rule: 4 mm Armox is treated as Ramor 500.
  $A4Armox=($A -match 'ARMOX' -and $A -match '4\s*MM')
  $B4Armox=($B -match 'ARMOX' -and $B -match '4\s*MM')
  $ARamor500=($A -match 'RAMOR' -and $A -match '500')
  $BRamor500=($B -match 'RAMOR' -and $B -match '500')
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
function Scan-Library([string]$root) {
  if(-not(Test-Path -LiteralPath $root)){throw "Cannot access the geometry library folder: $root"}
  if(-not((Get-Item -LiteralPath $root).PSIsContainer)){throw "Geometry library path is not a folder: $root"}

  $files=@();$scanErrors=@()
  try{
    # Search the configured server folder recursively. We index BOTH legacy
    # SigmaNEST .PRS files and new .DXF geometry files.
    $files=@(
      Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -ieq '.prs' -or $_.Extension -ieq '.dxf' } |
        ForEach-Object { $_.FullName }
    )
  }catch{
    $scanErrors += [pscustomobject]@{path=$root;error=$_.Exception.Message}
  }

  $items=@($files | ForEach-Object {
    $ext=[IO.Path]::GetExtension($_)
    $type=if($ext -ieq '.prs'){'PRS'}else{'DXF'}
    [pscustomobject]@{
      file=$_
      fileName=[IO.Path]::GetFileName($_)
      partName=[IO.Path]::GetFileNameWithoutExtension($_)
      embeddedPartName=[IO.Path]::GetFileNameWithoutExtension($_)
      fileType=$type
      likelyMaterial=''
      thickness=''
      sourceDxf=$(if($type -eq 'DXF'){$_}else{''})
      rotations=''
      metadataLoaded=($type -eq 'DXF')
    }
  })

  $script:INDEX=$items
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

  $prsCount=@($items|Where-Object {$_.fileType -eq 'PRS'}).Count
  $dxfCount=@($items|Where-Object {$_.fileType -eq 'DXF'}).Count

  $CFG.libraryRoot=$root
  $CFG.lastScan=(Get-Date).ToUniversalTime().ToString('o')
  $CFG.count=$items.Count
  $CFG.discoveredFiles=$files.Count
  $CFG.prsCount=$prsCount
  $CFG.dxfCount=$dxfCount
  $CFG.scanErrors=$scanErrors
  $CFG.inspectErrors=@()
  try{($CFG|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath $CFG_FILE -Encoding UTF8}catch{}

  [pscustomobject]@{
    count=$items.Count
    discoveredFiles=$files.Count
    prsCount=$prsCount
    dxfCount=$dxfCount
    scanErrors=$scanErrors
    inspectErrors=@()
    root=$root
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

  if($script:BYNAME -and $script:BYNAME.ContainsKey($n)){
    $hit=Select-MatchCandidate $script:BYNAME[$n]
    if($hit -and -not($hit.PSObject.Properties.Name -contains 'ambiguous')){
      Ensure-Metadata -item $hit | Out-Null
      Set-Prop $hit 'matchType' 'EXACT' | Out-Null
      return $hit
    }
    if($hit){return $hit}
  }

  # Embedded-name match across ALL subfolders and both source types.
  $embedded=@($script:INDEX|Where-Object {
    $en=Normalize -s $_.embeddedPartName
    $en -eq $n -or $en.StartsWith($n+'-')
  })
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
  if($script:BYVAR -and $script:BYVAR.ContainsKey($vk)){
    $vars=@($script:BYVAR[$vk])
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
    return [pscustomobject]@{Status=200;Data=@{ok=$true;port=$PORT;bridgeVersion=$BRIDGE_VERSION;libraryRoot=$CFG.libraryRoot;lastScan=$CFG.lastScan;count=[int]$CFG.count;discoveredFiles=[int]$CFG.discoveredFiles;prsCount=[int]$CFG.prsCount;dxfCount=[int]$CFG.dxfCount;sigmaNestCom=$true;bridge='PowerShell'}}
  }
  if($req.Path -eq '/api/scan' -and $req.Method -eq 'POST'){
    $b=if($req.Body){$req.Body|ConvertFrom-Json}else{[pscustomobject]@{}}
    $root=[IO.Path]::GetFullPath(([string]$(if($b.root){$b.root}else{$DEFAULT_LIBRARY})).Trim())
    $diag=Scan-Library -root $root
    return [pscustomobject]@{Status=200;Data=@{count=$diag.count;discoveredFiles=$diag.discoveredFiles;prsCount=$diag.prsCount;dxfCount=$diag.dxfCount;scanErrors=$diag.scanErrors;inspectErrors=$diag.inspectErrors;root=$diag.root;parts=@($script:INDEX|ForEach-Object{[pscustomobject]@{partName=$_.partName;embeddedPartName=$_.embeddedPartName;fileType=$_.fileType;file=$_.file;likelyMaterial=$_.likelyMaterial;thickness=$_.thickness;sourceDxf=$_.sourceDxf}})}}
  }
  if($req.Path -eq '/api/build-job' -and $req.Method -eq 'POST'){
    $b=$req.Body|ConvertFrom-Json
    $root=[IO.Path]::GetFullPath(([string]$(if($b.libraryRoot){$b.libraryRoot}else{$DEFAULT_LIBRARY})).Trim())
    if(-not $script:INDEX -or $script:INDEX.Count -eq 0 -or [string]$CFG.libraryRoot -ne [string]$root){
      Scan-Library -root $root|Out-Null
    }
    $name=([string]$(if($b.jobName){$b.jobName}else{'CL_JOB'}) -replace '[^A-Za-z0-9._ -]','_').Trim();if(-not$name){$name='CL_JOB'}
    $parts=@()
    foreach($p in @($b.parts)){
      $f=Find-Part -part ([string]$p.part)
      if(-not$f){
        Set-Prop $p 'status' 'MISSING' | Out-Null
        Set-Prop $p 'statusLabel' 'GEOMETRY MISSING' | Out-Null
        Set-Prop $p 'reviewReason' 'No matching .PRS or .DXF was found in the geometry library' | Out-Null
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
      if($matKnown -and $clMatKnown){$mok=Material-Equal -a $clMat -b $libMat}
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
    $request=[pscustomobject]@{jobName=$name;libraryRoot=$root;wsDirectory=[string]$b.wsDirectory;parts=@($parts|ForEach-Object{[pscustomobject]@{part=$_.part;qty=$_.qty;batchMultiplier=$(if($_.batchMultiplier){$_.batchMultiplier}else{1});taskSheet=$_.sheet;sourcePath=$_.sourcePath;sourceType=$_.sourceType;prsPath=$(if($_.sourceType -eq 'PRS'){[string]$_.sourcePath}else{''});sigmaMaterial=(Sigma-Material -cl ([string]$_.material) -lib ([string]$_.libraryMaterial));thicknessMm=(Thickness-Number -s ([string]$_.thickness))}})}
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
