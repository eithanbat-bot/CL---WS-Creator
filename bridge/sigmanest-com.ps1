# Pure SigmaNEST COM function library. Loaded by server.ps1, which runs in STA.
# No top-level execution and no exit statements.

function SN-Invoke-ComMethod($obj,[string]$name,[object[]]$args=@()){
  # Use reflection only for calls whose COM signature is not fixed in our
  # verified SigmaNEST interface. Verified SNApp methods are called directly
  # elsewhere below so PowerShell's COM binder supplies the correct signature.
  $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::InvokeMethod,$null,$obj,$args)
}

function SN-ErrorText($err){
  try{
    $ex=$err.Exception
    if($ex.InnerException){return [string]$ex.InnerException.Message}
    return [string]$ex.Message
  }catch{return [string]$err}
}
function SN-IsDisconnected($err){
  try{
    $hex=''
    if($err.Exception -and $err.Exception.HResult -ne $null){
      $hex=([int]$err.Exception.HResult).ToString('X8')
    }
    if($hex -eq '80010108'){return $true}
    $msg=[string]$err.Exception.Message
    return ($msg -match '(?i)0x80010108|RPC_E_DISCONNECTED|disconnected from its clients')
  }catch{return $false}
}
function SN-Try-Set($obj,[string[]]$names,$value){
  if($null -eq $obj){return $null}
  foreach($name in $names){
    if([string]::IsNullOrWhiteSpace($name)){continue}
    try{$obj.$name=$value;return $name}catch{}
    try{$obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::SetProperty,$null,$obj,@($value))|Out-Null;return $name}catch{}
  }
  return $null
}

function SN-Set-Required($obj,[string[]]$names,$value,[string]$label){
  $set=SN-Try-Set -obj $obj -names $names -value $value
  if(-not $set){throw ('SigmaNEST part does not expose a writable '+$label+' property. Tried: '+($names -join ', '))}
  return $set
}

function SN-Safe-Set($obj,[string]$name,$value){[void](SN-Try-Set -obj $obj -names @($name) -value $value)}

function SN-Scalar-Number($value,[object]$default=1){
  try{
    if($null -eq $value){return $default}
    if($value -is [System.Array]){
      $vals=@($value|Where-Object {$_ -ne $null})
      if($vals.Count -eq 0){return $default}
      $value=$vals[0]
    }
    if($value -is [string]){
      $s=$value.Trim()
      if([string]::IsNullOrWhiteSpace($s) -or $s -match '(?i)^(?:\[double\]::)?(?:NaN|Infinity|-Infinity)$'){return $default}
    }
    $d=[double]$value
    if([double]::IsNaN($d) -or [double]::IsInfinity($d)){return $default}
    return $d
  }catch{return $default}
}
function SN-Scalar-Int($value,$default=1){
  $n=SN-Scalar-Number -value $value -default $default
  try{return [int][math]::Round([double]$n)}catch{
    try{return [int]$default}catch{return 1}
  }
}
function SN-Get-WorkspacePartByIndex($app,$index){
  $index=SN-Scalar-Int -value $index -default -1;
  if($index -lt 0){return $null}
  try{
    $count=SN-Parts-Count $app
    if($index -ge $count){return $null}
    $part=SN-Invoke-ComMethod -obj $app.PartsList -name 'Items' -args @([int]$index)
    if($null -ne $part){return [pscustomobject]@{part=$part;index=$index}}
  }catch{}
  return $null
}
function SN-Find-WorkspacePart($app,[string]$targetName){
  $count=SN-Parts-Count $app
  for($i=0;$i -lt $count;$i++){
    $part=$null
    try{$part=$app.PartsList.Items($i)}catch{continue}
    if($null -eq $part){continue}
    $name=''
    try{$name=[string]$part.Name}catch{}
    if(-not [string]::IsNullOrWhiteSpace($name) -and $name.Equals($targetName,[StringComparison]::OrdinalIgnoreCase)){
      return [pscustomobject]@{part=$part;index=$i}
    }
  }
  return $null
}
function SN-Get-NestedObject($obj,[string[]]$names){
  if($null -eq $obj){return @()}
  $out=@()
  foreach($name in $names){
    try{
      $v=$obj.$name
      if($null -ne $v){$out+=[pscustomobject]@{name=$name;object=$v}}
    }catch{}
  }
  @($out)
}
function SN-Try-SetField($obj,[string[]]$names,$value,[string]$expectedText='',$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $obj){return $null}
  foreach($name in $names){
    try{
      $obj.$name=$value
      if($expectedText -ne ''){
        if(([string]$obj.$name).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $name}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        if(([double]$obj.$name) -eq $expectedNumber){return $name}
      }elseif($expectedInt -ne -2147483648){
        if(([int]$obj.$name) -eq $expectedInt){return $name}
      }else{return $name}
    }catch{}
    try{
      $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::SetProperty,$null,$obj,@($value))|Out-Null
      if($expectedText -ne ''){
        $back=[string]$obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())
        if($back.Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $name}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        try{if(([double]$obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())) -eq $expectedNumber){return $name}}catch{}
      }elseif($expectedInt -ne -2147483648){
        try{if(([int]$obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())) -eq $expectedInt){return $name}}catch{}
      }else{return $name}
    }catch{}
  }
  return $null
}
function SN-ComPropertyNames($obj){
  if($null -eq $obj){return @()}
  try{
    return @($obj | Get-Member -MemberType Property -ErrorAction Stop | Select-Object -ExpandProperty Name -Unique)
  }catch{return @()}
}
function SN-ComMethodNames($obj){
  if($null -eq $obj){return @()}
  try{
    return @($obj | Get-Member -MemberType Method -ErrorAction Stop | Select-Object -ExpandProperty Name -Unique)
  }catch{return @()}
}
function SN-NormalizeFieldName([string]$name){
  if($null -eq $name){return ''}
  ($name -replace '[^A-Za-z0-9]','').ToUpperInvariant()
}
function SN-FieldNameMatches([string]$name,[string[]]$aliases){
  $n=SN-NormalizeFieldName $name
  foreach($a in @($aliases)){
    if($n -eq (SN-NormalizeFieldName $a)){return $true}
  }
  return $false
}
function SN-Try-InvokeGetter($obj,[string]$name){
  if($null -eq $obj -or [string]::IsNullOrWhiteSpace($name)){return $null}
  try{
    $m=$obj | Get-Member -MemberType Method -Name $name -ErrorAction Stop | Select-Object -First 1
    if($m){
      $method=$m
      $params=@()
      if($method.Definition -notmatch '\([^)]*\)'){return $null}
      if($method.Definition -match '\([^)]*[^\s()]([^)]*)\)' -and $Matches[1].Trim()){return $null}
      return $obj.$name()
    }
  }catch{}
  return $null
}
function SN-Try-SetDiscoveredField($obj,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $obj){return $null}
  foreach($prop in @(SN-ComPropertyNames $obj)){
    if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
    try{
      $obj.$prop=$value
      if($expectedText -ne ''){
        if(([string]$obj.$prop).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        $back=SN-Scalar-Number -value $obj.$prop -default ([double]::NaN)
        if(-not [double]::IsNaN([double]$back) -and [double]$back -eq $expectedNumber){return $prop}
      }elseif($expectedInt -ne -2147483648){
        $back=SN-Scalar-Int -value $obj.$prop -default -2147483648
        if($back -eq $expectedInt){return $prop}
      }else{return $prop}
    }catch{}
    try{
      $obj.GetType().InvokeMember($prop,[Reflection.BindingFlags]::SetProperty,$null,$obj,@($value))|Out-Null
      if($expectedText -ne ''){
        $back=[string]$obj.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())
        if($back.Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        $back=SN-Scalar-Number -value $obj.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$obj,@()) -default ([double]::NaN)
        if(-not [double]::IsNaN([double]$back) -and [double]$back -eq $expectedNumber){return $prop}
      }elseif($expectedInt -ne -2147483648){
        $back=SN-Scalar-Int -value $obj.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$obj,@()) -default -2147483648
        if($back -eq $expectedInt){return $prop}
      }else{return $prop}
    }catch{}
  }
  return $null
}
function SN-Set-PartFieldDiscovered($partObj,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $partObj){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$partObj;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 4){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{$identity=''}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    $hit=SN-Try-SetDiscoveredField -obj $obj -aliases $aliases -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
    if($hit){return [pscustomobject]@{path=($node.path+'.'+$hit);object=$obj}}

    if($depth -ge 4){continue}
    foreach($prop in @(SN-ComPropertyNames $obj)){
      if($prop -in @('OwnerList','ParentObject','Parent','Application','Owner','Geometry','PartPolyLinesList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
    foreach($method in @(SN-ComMethodNames $obj | Where-Object {$_ -match '^(Get|Fetch|Read).*(Part|Param|Data|Record|Info)' -and $_ -notmatch '(Count|List|Item|Current|Owner)' })){
      try{
        $child=SN-Try-InvokeGetter -obj $obj -name $method
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$method+'()');depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Set-PartField($partObj,[string[]]$names,$value,[string]$expectedText='',$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  $set=SN-Try-SetField -obj $partObj -names $names -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
  if($set){return [pscustomobject]@{path=$set;object=$partObj}}
  $children=@(SN-Get-NestedObject -obj $partObj -names @('PNVar','PartData','PartParameters','PartParameter','Parameters','Param','PartRec','PartInfo','PartParameterData','PartRecord','PartParams','ParametersData'))
  foreach($child in $children){
    $nested=SN-Try-SetField -obj $child.object -names $names -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
    if($nested){return [pscustomobject]@{path=($child.name+'.'+$nested);object=$child.object}}
    foreach($grand in SN-Get-NestedObject -obj $child.object -names @('PNVar','PartData','PartParameters','PartParameter','Parameters','Param','PartRec','PartInfo','PartParameterData','PartRecord','PartParams','ParametersData')){
      $nested2=SN-Try-SetField -obj $grand.object -names $names -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
      if($nested2){return [pscustomobject]@{path=($child.name+'.'+$grand.name+'.'+$nested2);object=$grand.object}}
    }
  }
  return (SN-Set-PartFieldDiscovered -partObj $partObj -aliases $names -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt)
}
function SN-Normalize-PartIdentity([string]$value){
  if($null -eq $value){return ''}
  return (($value.ToUpperInvariant() -replace '\\.[Pp][Rr][Ss]$','') -replace '[^A-Z0-9]','')
}
function SN-Get-PartIdentity($partObj){
  if($null -eq $partObj){return [pscustomobject]@{name='';drawing='';file='';normalizedName='';normalizedDrawing='';normalizedFile='';partNumber='';normalizedPartNumber=''}}
  $name='';$drawing='';$file='';$partNumber=''
  try{$name=[string]$partObj.Name}catch{}
  foreach($prop in @('DrawingNumber','DrawingNo','DwgNumber','DwgNo')){
    if([string]::IsNullOrWhiteSpace($drawing)){try{$drawing=[string]$partObj.$prop}catch{}}
  }
  foreach($prop in @('PartFilename','SourceFilePath','SourcePath','PartFileName','Filename','FileName')){
    if([string]::IsNullOrWhiteSpace($file)){try{$file=[string]$partObj.$prop}catch{}}
  }
  foreach($prop in @('PartNumber','PartNo','Number')){
    if([string]::IsNullOrWhiteSpace($partNumber)){try{$partNumber=[string]$partObj.$prop}catch{}}
  }
  return [pscustomobject]@{
    name=$name
    drawing=$drawing
    file=$file
    partNumber=$partNumber
    normalizedName=(SN-Normalize-PartIdentity $name)
    normalizedDrawing=(SN-Normalize-PartIdentity $drawing)
    normalizedFile=(SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension($file)))
    normalizedPartNumber=(SN-Normalize-PartIdentity $partNumber)
  }
}
function SN-RequestSourceStem($rp){
  $source=''
  try{$source=[string]$rp.sourcePath}catch{}
  if([string]::IsNullOrWhiteSpace($source)){try{$source=[string]$rp.prsPath}catch{}}
  if([string]::IsNullOrWhiteSpace($source)){return ''}
  $stem=''
  try{$stem=[IO.Path]::GetFileNameWithoutExtension($source)}catch{}
  $stem=(SN-Normalize-PartIdentity $stem)
  return (($stem -replace '(?i)PRSfunction SN-Find-WorkspacePartExact($app,$targetName,$sourcePath='',$usedIndices=@()){
  $targetNorm=SN-Normalize-PartIdentity $targetName
  $sourceStem=''
  try{$sourceStem=SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension([string]$sourcePath))}catch{}
  $sourceStemNoPrs=SN-Normalize-PartIdentity ($sourceStem -replace '(?i)PRS$','')
  $count=SN-Parts-Count $app
  $candidates=@()
  for($i=0;$i -lt $count;$i++){
    if(@($usedIndices) -contains $i){continue}
    $part=$null
    try{$part=$app.PartsList.Items($i)}catch{continue}
    if($null -eq $part){continue}
    $id=SN-Get-PartIdentity $part
    $score=0
    if($targetNorm -and $id.normalizedName -eq $targetNorm){$score=100}
    if($sourceStemNoPrs -and $id.normalizedFile -eq $sourceStemNoPrs){$score=[math]::Max($score,95)}
    if($targetNorm -and $id.normalizedDrawing -eq $targetNorm){$score=[math]::Max($score,90)}
    if($sourceStemNoPrs -and $id.normalizedDrawing -eq $sourceStemNoPrs){$score=[math]::Max($score,85)}
    if($score -gt 0){
      $candidates += [pscustomobject]@{part=$part;index=$i;score=$score;identity=$id}
    }
  }
  if($candidates.Count -eq 0){return $null}
  $top=@($candidates|Sort-Object @{Expression={[int]$_.score};Descending=$true},@{Expression={[int]$_.index};Descending=$false})
  if($top.Count -gt 1 -and [int]$top[0].score -eq [int]$top[1].score){
    # Prefer an exact normalized Name match over filename/drawing fallbacks.
    $exact=@($top|Where-Object {$_.identity.normalizedName -eq $targetNorm})
    if($exact.Count -eq 1){return $exact[0]}
    if($exact.Count -gt 1){throw ('Multiple SigmaNEST workspace parts match CL part "'+$targetName+'".')}
  }
  return $top[0]
}
function SN-Verify-PartIdentity($partObj,$targetName,$sourcePath=''){
  $id=SN-Get-PartIdentity $partObj
  $targetNorm=SN-Normalize-PartIdentity $targetName
  $sourceNorm=''
  try{$sourceNorm=SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension([string]$sourcePath))}catch{}
  $sourceNorm=SN-Normalize-PartIdentity ($sourceNorm -replace '(?i)PRS$','')
  if($targetNorm -and $id.normalizedName -eq $targetNorm){return $true}
  if($targetNorm -and $id.normalizedDrawing -eq $targetNorm){return $true}
  if($sourceNorm -and $id.normalizedFile -eq $sourceNorm){return $true}
  if($sourceNorm -and $id.normalizedDrawing -eq $sourceNorm){return $true}
  return $false
}
function SN-Make-CLLinkKey($jobName,$rp){
  $source=SN-RequestSourceStem $rp
  $raw=([string]$jobName+'|'+[string]$rp.part+'|'+$source)
  try{
    $sha=[Security.Cryptography.SHA256]::Create()
    $bytes=[Text.Encoding]::UTF8.GetBytes($raw)
    $hash=$sha.ComputeHash($bytes)
    return ([BitConverter]::ToString($hash) -replace '-','').Substring(0,16)
  }catch{
    return (($raw -replace '[^A-Za-z0-9]+','_').Trim('_')).Substring(0,[math]::Min(40,(($raw -replace '[^A-Za-z0-9]+','_').Trim('_')).Length))
  }
}
function SN-Apply-WorkspacePartData($app,$requestParts,[string]$jobName='',$linkFile=''){
  $updated=@()
  $usedIndices=@()
  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $sourcePath=[string]$rp.sourcePath
    $found=SN-Find-WorkspacePartExact -app $app -targetName $targetName -sourcePath $sourcePath -usedIndices $usedIndices
    if($null -eq $found){
      throw ('Could not map imported SigmaNEST geometry back to CL part "'+$targetName+'". The imported PartsList order is not trusted.')
    }
    if(-not (SN-Verify-PartIdentity -partObj $found.part -targetName $targetName -sourcePath $sourcePath)){
      throw ('SigmaNEST workspace identity verification failed for CL part "'+$targetName+'". Refusing to overwrite another part.')
    }
    $usedIndices += [int]$found.index
    $partObj=$found.part
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $material=[string]$rp.sigmaMaterial
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    $linkKey=SN-Make-CLLinkKey -jobName $jobName -rp $rp
    $row=[ordered]@{
      linkKey=$linkKey
      part=$targetName
      index=[int]$found.index
      qty=$qty
      material=$material
      thickness=$thickness
      sourcePath=$sourcePath
      sourceType=[string]$rp.sourceType
      quantityProperty=''
      materialProperty=''
      thicknessProperty=''
      drawingProperty=''
      workOrderProperty=''
      warnings=@()
    }

    if(-not [string]::IsNullOrWhiteSpace($material)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @(
        'Material','MaterialName','PartMaterial','MaterialDescription','MaterialType',
        'Mat','MatName','SheetMaterial','StockMaterial','NestMaterial'
      ) -value $material -expectedText $material
      if(-not $setInfo){
        throw ('CL material "'+$material+'" could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.materialProperty=[string]$setInfo.path
    }

    if(-not [double]::IsNaN($thickness)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @(
        'Thickness','SheetThickness','MaterialThickness','ThicknessValue','PartThickness','Thk','Thick'
      ) -value $thickness -expectedNumber $thickness
      if(-not $setInfo){
        throw ('CL thickness '+$thickness+'mm could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.thicknessProperty=[string]$setInfo.path
    }

    $setInfo=SN-Set-PartField -partObj $partObj -names @(
      'QtyOrdered','Quantity','PartQuantity','QuantityOrdered','QtyRequired','QtyReq',
      'BatchQty','BatchQuantity','QtyToNest','QuantityToNest','NestQuantity'
    ) -value $qty -expectedInt $qty
    if(-not $setInfo){
      throw ('CL quantity '+$qty+' could not be written and verified on SigmaNEST part "'+$targetName+'".')
    }
    $row.quantityProperty=[string]$setInfo.path

    # These visible identity fields create a durable CL -> SigmaNEST association.
    # Do not rename geometry: keep the native SigmaNEST name, but put the CL drawing
    # number and job number into standard part fields when those fields exist.
    foreach($item in @(
      @{aliases=@('DrawingNumber','DrawingNo','DwgNumber','DwgNo','PartNumber','PartNo');value=$targetName;label='drawingProperty'},
      @{aliases=@('WONumber','WorkOrder','WorkOrderNumber','OrderNumber');value=[string]$jobName;label='workOrderProperty'}
    )){
      $setInfo=SN-Set-PartField -partObj $partObj -names $item.aliases -value ([string]$item.value) -expectedText ([string]$item.value)
      if($setInfo){$row[$item.label]=[string]$setInfo.path}
    }

    $updated += [pscustomobject]$row
  }

  if(-not [string]::IsNullOrWhiteSpace($linkFile)){
    $linkRows=@()
    foreach($rp in @($requestParts)){
      $match=@($updated|Where-Object {$_.part -eq [string]$rp.part}|Select-Object -First 1)
      $linkRows += [pscustomobject]@{
        linkKey=$(if($match.Count){[string]$match[0].linkKey}else{SN-Make-CLLinkKey -jobName $jobName -rp $rp})
        jobName=[string]$jobName
        clPart=[string]$rp.part
        clMaterial=[string]$rp.material
        sigmaMaterial=[string]$rp.sigmaMaterial
        clThickness=[string]$rp.thicknessMm
        clQuantity=[SN-Scalar-Int -value $rp.qty -default 1]
        batchMultiplier=[SN-Scalar-Int -value $rp.batchMultiplier -default 1]
        sourceType=[string]$rp.sourceType
        sourcePath=[string]$rp.sourcePath
        sourceSheets=@($rp.sourceSheets)
        sourceRows=@($rp.sourceRows)
        sigmaNestPartIndex=$(if($match.Count){[int]$match[0].index}else{-1})
      }
    }
    try{
      $parent=Split-Path -Parent $linkFile
      if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
      ([ordered]@{
        schema='cl-ws-creator/cl-data-link/1.0'
        jobName=[string]$jobName
        createdUtc=(Get-Date).ToUniversalTime().ToString('o')
        records=$linkRows
      }|ConvertTo-Json -Depth 20)|Set-Content -LiteralPath $linkFile -Encoding UTF8
    }catch{
      # Link sidecar is valuable but must never silently make a correct
      # SigmaNEST import fail.
    }
  }

  @($updated)
}
function SN-Read-PartField($partObj,[string[]]$aliases,[string]$expectedText='',$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $partObj){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$partObj;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 4){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    foreach($prop in @(SN-ComPropertyNames $obj)){
      if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
      try{
        $v=$obj.$prop
        if($expectedText -ne ''){
          if(([string]$v).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=[string]$v}
          }
        }elseif(-not [double]::IsNaN($expectedNumber)){
          $n=SN-Scalar-Number -value $v -default ([double]::NaN)
          if(-not [double]::IsNaN([double]$n) -and [double]$n -eq $expectedNumber){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=$n}
          }
        }elseif($expectedInt -ne -2147483648){
          $n=SN-Scalar-Int -value $v -default -2147483648
          if($n -eq $expectedInt){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=$n}
          }
        }else{
          return [pscustomobject]@{path=($node.path+'.'+$prop);value=$v}
        }
      }catch{}
    }

    if($depth -ge 4){continue}
    foreach($prop in @(SN-ComPropertyNames $obj)){
      if($prop -in @('OwnerList','ParentObject','Parent','Application','Owner','Geometry','PartPolyLinesList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Verify-WorkspaceCLData($app,$requestParts){
  foreach($rp in @($requestParts)){
    $target=[string]$rp.part
    $source=[string]$rp.sourcePath
    $found=SN-Find-WorkspacePartExact -app $app -targetName $target -sourcePath $source -usedIndices @()
    if($null -eq $found){throw ('Post-save verification could not find SigmaNEST part "'+$target+'".')}
    $material=[string]$rp.sigmaMaterial
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','PartMaterial','MaterialDescription','MaterialType','Mat','MatName','SheetMaterial','StockMaterial','NestMaterial') -expectedText $material
      if($null -eq $m){throw ('Post-save verification failed for "'+$target+'": material is not "'+$material+'".')}
    }
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    if(-not [double]::IsNaN($thickness)){
      $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','MaterialThickness','ThicknessValue','PartThickness','Thk','Thick') -expectedNumber $thickness
      if($null -eq $t){throw ('Post-save verification failed for "'+$target+'": thickness is not '+$thickness+'mm.')}
    }
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $q=SN-Read-PartField -partObj $found.part -aliases @('QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity','BatchQty','BatchQuantity','QtyToNest','QuantityToNest','NestQuantity') -expectedInt $qty
    if($null -eq $q){throw ('Post-save verification failed for "'+$target+'": quantity is not '+$qty+'.')}
  }
  return $true
}
function SN-Parts-Count($app){try{return [int]$app.PartsList.Count}catch{return 0}}

function SN-Get-NewPart($app,$beforeCount,[string]$label){
  $beforeCount=SN-Scalar-Int -value $beforeCount -default 0;
  $afterCount=SN-Parts-Count $app
  if($afterCount -le $beforeCount){throw ('SigmaNEST loaded "'+$label+'" but did not add it to the workspace PartsList.')}
  try{$last=$null;foreach($item in $app.PartsList){$last=$item};if($null -ne $last){return $last}}catch{}
  foreach($member in @('Item','Items','get_Item')){
    try{$obj=$app.PartsList.GetType().InvokeMember($member,[Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::GetProperty,$null,$app.PartsList,@([int]($afterCount-1)));if($null -ne $obj){return $obj}}catch{}
  }
  throw ('SigmaNEST added "'+$label+'" but the new PartsList item could not be accessed. PartsList.Count='+$afterCount)
}

function SN-AllMemberDefinitions($obj){
  if($null -eq $obj){return @()}
  try{
    return @($obj | Get-Member -MemberType Methods,Properties,ParameterizedProperty,CodeProperty,NoteProperty,ScriptProperty,Fields -ErrorAction Stop |
      ForEach-Object {[pscustomobject]@{Name=[string]$_.Name;Type=[string]$_.MemberType;Definition=[string]$_.Definition}})
  }catch{return @()}
}
function SN-ComPropertyNamesAll($obj){
  return @((SN-AllMemberDefinitions $obj)|Where-Object {$_.Type -match 'Property'}|Select-Object -ExpandProperty Name -Unique)
}
function SN-InvokeSingleArgMember($obj,[string]$methodName,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $obj -or [string]::IsNullOrWhiteSpace($methodName)){return $null}
  try{
    $defs=@((SN-AllMemberDefinitions $obj)|Where-Object {$_.Name -eq $methodName})
    foreach($d in $defs){
      $def=[string]$d.Definition
      if($def -notmatch '\\(([^)]*)\\)'){continue}
      $inside=$Matches[1].Trim()
      if($inside -match ','){continue}
      if($inside -and $inside -notmatch '(?i)optional|paramarray'){ }
      try{
        $result=$obj.GetType().InvokeMember($methodName,[Reflection.BindingFlags]::InvokeMethod,$null,$obj,@($value))
        if($expectedText -ne ''){
          if($result -ne $null -and ([string]$result).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $methodName}
          try{
            $back=$obj.GetType().InvokeMember($methodName,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())
            if(([string]$back).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $methodName}
          }catch{}
        }elseif(-not [double]::IsNaN($expectedNumber)){
          try{if(([double]$result) -eq $expectedNumber){return $methodName}}catch{}
        }elseif($expectedInt -ne -2147483648){
          try{if((SN-Scalar-Int $result -default -2147483648) -eq $expectedInt){return $methodName}}catch{}
        }else{return $methodName}
      }catch{}
    }
  }catch{}
  return $null
}
function SN-Try-SetImportSetting($settings,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $settings){return $null}
  # 1) Standard writable/parameterized properties.
  foreach($prop in @(SN-ComPropertyNamesAll $settings)){
    if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
    try{
      $settings.$prop=$value
      if($expectedText -ne '' -and ([string]$settings.$prop).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      if(-not [double]::IsNaN($expectedNumber) -and (SN-Scalar-Number $settings.$prop -default ([double]::NaN)) -eq $expectedNumber){return $prop}
      if($expectedInt -ne -2147483648 -and (SN-Scalar-Int $settings.$prop -default -2147483648) -eq $expectedInt){return $prop}
    }catch{}
    try{
      $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::SetProperty,$null,$settings,@($value))|Out-Null
      if($expectedText -ne ''){
        $back=[string]$settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@())
        if($back.Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        $back=SN-Scalar-Number $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@()) -default ([double]::NaN)
        if(-not [double]::IsNaN($back) -and $back -eq $expectedNumber){return $prop}
      }elseif($expectedInt -ne -2147483648){
        $back=SN-Scalar-Int $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@()) -default -2147483648
        if($back -eq $expectedInt){return $prop}
      }
    }catch{}
  }

  # 2) Setter-style methods such as SetMaterial / SetThickness / SetQuantity.
  foreach($method in @(SN-AllMemberDefinitions $settings | Where-Object {$_.Type -eq 'Method'} | Select-Object -ExpandProperty Name -Unique)){
    $norm=SN-NormalizeFieldName $method
    foreach($alias in @($aliases)){
      $an=SN-NormalizeFieldName $alias
      if($norm -eq $an -or $norm -eq ('SET'+$an) -or $norm -eq ('SET'+$an+'VALUE') -or $norm -eq ('SET'+$an+'NAME') -or ($norm.StartsWith('SET') -and $norm.Contains($an))){
        $hit=SN-InvokeSingleArgMember -obj $settings -methodName $method -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
        if($hit){return ('METHOD.'+$hit)}
      }
    }
  }

  # 3) Generic field/value methods used by COM import-mapping objects.
  foreach($method in @('SetValue','SetField','SetParameter','SetProperty','SetImportValue','SetOption')){
    $defs=@((SN-AllMemberDefinitions $settings)|Where-Object {$_.Name -eq $method -and $_.Type -eq 'Method'})
    foreach($d in $defs){
      try{
        if(([string]$d.Definition) -match '\\([^)]*\\)' -and [string]$d.Definition -notmatch ','){
          $r=$settings.GetType().InvokeMember($method,[Reflection.BindingFlags]::InvokeMethod,$null,$settings,@($value))
          if($expectedText -ne '' -and $r -ne $null -and ([string]$r).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return ('GENERIC.'+$method)}
          if($expectedInt -ne -2147483648 -and (SN-Scalar-Int $r -default -2147483648) -eq $expectedInt){return ('GENERIC.'+$method)}
          if(-not [double]::IsNaN($expectedNumber) -and (SN-Scalar-Number $r -default ([double]::NaN)) -eq $expectedNumber){return ('GENERIC.'+$method)}
        }
      }catch{}
    }
  }
  return $null
}
function SN-Get-ImportSettings($adapter){
  if($null -ne $adapter){
    foreach($prop in @('PartImportSettings','ImportSettings','Settings')){
      try{
        $v=$adapter.$prop
        if($null -ne $v){return [pscustomobject]@{settings=$v;property=$prop}}
      }catch{}
    }
  }
  foreach($progId in @('SigmaNEST.SNPartImportSettings','SigmaNEST.SNPartImportSetting')){
    try{
      $v=New-Object -ComObject $progId -ErrorAction Stop
      if($null -ne $v){return [pscustomobject]@{settings=$v;property=$progId}}
    }catch{}
  }
  return $null
}
function SN-Get-ImportDiagnostics($adapter,$settings){
  $parts=@()
  if($null -ne $adapter){
    $parts+='Adapter methods/properties:'
    foreach($d in @(SN-AllMemberDefinitions $adapter)){$parts+=('  '+$d.Name+' :: '+$d.Definition)}
  }
  if($null -ne $settings){
    $parts+='PartImportSettings methods/properties:'
    foreach($d in @(SN-AllMemberDefinitions $settings)){$parts+=('  '+$d.Name+' :: '+$d.Definition)}
    foreach($prop in @(SN-ComPropertyNamesAll $settings)){
      try{
        $child=$settings.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array] -and $child.Count -gt 1){continue}
        $parts+=('  Nested '+$prop+':')
        foreach($d in @(SN-AllMemberDefinitions $child)){$parts+=('    '+$d.Name+' :: '+$d.Definition)}
      }catch{}
    }
  }
  return @($parts)
}
function SN-SetImportSettingDeep($root,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $root){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$root;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 5){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    $hit=SN-Try-SetImportSetting -settings $obj -aliases $aliases -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
    if($hit){return [pscustomobject]@{path=($node.path+'.'+$hit)}}

    if($depth -ge 5){continue}
    foreach($prop in @(SN-ComPropertyNamesAll $obj)){
      if($prop -in @('Owner','Parent','ParentObject','Application','PartImportSettings','SolidsList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Configure-PRS-ImportSettings($settings,$clData){
  $material=[string]$clData.sigmaMaterial
  $thickness=SN-Scalar-Number -value $clData.thicknessMm -default ([double]::NaN)
  $qty=SN-Scalar-Int -value $clData.qty -default 1
  if($qty -lt 1){$qty=1}
  $result=[ordered]@{material='';thickness='';quantity='';warnings=@()}
  if(-not [string]::IsNullOrWhiteSpace($material)){
    $r=SN-SetImportSettingDeep -root $settings -aliases @('Material','MaterialName','PartMaterial','MaterialType','MaterialString') -value $material -expectedText $material
    if($r){$result.material=$r.path}
  }
  if(-not [double]::IsNaN($thickness)){
    $r=SN-SetImportSettingDeep -root $settings -aliases @('Thickness','SheetThickness','MaterialThickness','MaterialThk','Thk','Thick') -value $thickness -expectedNumber $thickness
    if($r){$result.thickness=$r.path}
  }
  $r=SN-SetImportSettingDeep -root $settings -aliases @('QtyOrdered','Quantity','PartQuantity','NumberToNest','NumberToLoad','QuantityToNest','QtyToNest','BatchQuantity') -value $qty -expectedInt $qty
  if($r){$result.quantity=$r.path}
  if([string]::IsNullOrWhiteSpace([string]$result.material)){ $result.warnings+='Import settings did not expose a verified material field.' }
  if(-not [double]::IsNaN($thickness) -and [string]::IsNullOrWhiteSpace([string]$result.thickness)){ $result.warnings+='Import settings did not expose a verified thickness field.' }
  if([string]::IsNullOrWhiteSpace([string]$result.quantity)){ $result.warnings+='Import settings did not expose a verified quantity field.' }
  return [pscustomobject]$result
}
function SN-Try-ImportPRS-WithSettings($app,[string]$sourcePath,$clData){
  $before=SN-Parts-Count $app
  $adapter=$null
  $info=$null
  $diagnostics=@()
  try{$adapter=New-Object -ComObject SigmaNEST.SNPartExportImport -ErrorAction Stop}catch{
    $diagnostics+='SNPartExportImport unavailable: '+(SN-ErrorText $_)
  }
  if($null -ne $adapter){
    $info=SN-Get-ImportSettings -adapter $adapter
    if($info){
      $cfg=SN-Configure-PRS-ImportSettings -settings $info.settings -clData $clData
      if($cfg.warnings.Count -eq 0){
        foreach($methodName in @('ImportPartWithFeedback','ImportPart2','ImportPart','ImportAsParts')){
          $members=@(SN-ComMethodNames $adapter)
          if($members -notcontains $methodName){continue}
          foreach($args in @(
            @([string]$sourcePath,$info.settings),
            @([string]$sourcePath,$info.settings,$false),
            @([string]$sourcePath)
          )){
            try{
              $result=SN-Invoke-ComMethod -obj $adapter -name $methodName -args $args
              Start-Sleep -Milliseconds 150
              $after=SN-Parts-Count $app
              if($after -gt $before){
                return [pscustomobject]@{ok=$true;method=('SNPartExportImport.'+$methodName);settings=$cfg}
              }
              if($result -is [System.Object[]]){
                $result=@($result|Where-Object {$null -ne $_})[0]
              }
              if($result -and $result.PSObject.Properties.Name -contains 'Part'){
                try{
                  $part=$result.Part
                  if($null -ne $part){return [pscustomobject]@{ok=$true;method=('SNPartExportImport.'+$methodName);settings=$cfg;pendingPart=$part}}
                }catch{}
              }
            }catch{
              $diagnostics+=($methodName+' '+$args.Count+' args: '+(SN-ErrorText $_))
            }
          }
        }
      }else{
        $diagnostics+=($cfg.warnings -join ' ')
      }
    }else{
      $diagnostics+='SNPartExportImport has no accessible PartImportSettings/ImportSettings/Settings object.'
    }
  }
  if($null -ne $adapter){
    try{$diagnostics+='Adapter methods: '+((SN-ComMethodNames $adapter) -join ', ')}catch{}
    try{$diagnostics+='Adapter properties: '+((SN-ComPropertyNames $adapter) -join ', ')}catch{}
  }
  try{if($adapter){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($adapter)}}catch{}
  if($null -ne $adapter -and $null -ne $info -and $null -ne $info.settings){
    $diagnostics += @(SN-Get-ImportDiagnostics -adapter $adapter -settings $info.settings)
  }
  return [pscustomobject]@{ok=$false;diagnostics=$diagnostics}
}
function SN-Queue-Geometry($app,[string]$sourcePath,[string]$sourceType,$clData=$null){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $errors=@()

  # PRS is geometry-only. LoadPart is the preferred loader because the
  # authoritative CL material/thickness/quantity is applied after the part is
  # committed to the workspace and verified there.
  if($sourceType -eq 'PRS'){
    try{
      [void]$app.LoadPart([string]$sourcePath)
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method='LoadPart-PRS-GEOMETRY'}
    }catch{
      $errors+=('LoadPart PRS: '+(SN-ErrorText $_))
    }

    # Fallback for installations where PRS must use the import/export adapter.
    # Even when this path succeeds, its original PRS metadata is NEVER trusted:
    # SN-Apply-WorkspacePartData overwrites and verifies the CL values.
    if($null -ne $clData){
      $prsTry=SN-Try-ImportPRS-WithSettings -app $app -sourcePath $sourcePath -clData $clData
      if($prsTry.ok){
        return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method=$prsTry.method;preImportSettings=$prsTry.settings}
      }
      if($prsTry.diagnostics.Count){$errors+=($prsTry.diagnostics -join ' | ')}
    }
  }else{
    try{
      [void]$app.LoadPart([string]$sourcePath)
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method='LoadPart'}
    }catch{
      $errors+=('LoadPart: '+(SN-ErrorText $_))
    }
  }

  if($sourceType -eq 'DXF'){
    try{
      $settings=$null;try{$settings=New-Object -ComObject SigmaNEST.SNPartImportSettings}catch{}
      foreach($importId in @(0,1,2,3)){
        try{
          $result=SN-Invoke-ComMethod $app.PartsList 'Import' @([string]$sourcePath,[int]$importId,$settings)
          return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method=('PartsList.Import '+$importId)}
        }catch{
          $errors+=('PartsList.Import id '+$importId+': '+(SN-ErrorText $_))
        }
      }
    }catch{
      $errors+=('DXF import fallback: '+(SN-ErrorText $_))
    }
  }

  throw ('SigmaNEST could not load "'+$label+'". Geometry source: '+$sourceType+'. '+($errors -join ' | '))
}


function SN-Set-TaskMaterialAndThickness($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){
    throw 'SigmaNEST created no TasksList entries after CreateTasksListForNewPartsInWS; cannot apply CL material/thickness safely.'
  }

  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $material=[string]$rp.sigmaMaterial
    $thickness=$null
    try{
      if($rp.thicknessMm -ne $null -and -not [double]::IsNaN([double]$rp.thicknessMm)){
        $thickness=[double]$rp.thicknessMm
      }
    }catch{}

    $matched=$false
    $matchedTaskIndex=-1
    $matchedPartIndex=-1

    for($ti=0;$ti -lt $taskCount -and -not $matched;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}

      $taskPartCount=0
      try{$taskPartCount=[int]$task.PartsList.Count}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}
        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not (SN-TaskPart-MatchesRequest -taskPart $taskPart -rp $rp)){continue}

        $matched=$true
        $matchedTaskIndex=$ti
        $matchedPartIndex=$pi
        break
      }
    }

    if(-not $matched){
      throw ('Could not find CL part "'+$targetName+'" inside any SigmaNEST task after task creation.')
    }

    # Reacquire the COM task/part immediately before each mutation. SigmaNEST
    # may rebuild the task tree when Task Setup data changes, invalidating
    # previously returned COM interfaces.
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $task=$null;$taskPart=$null
      try{$task=$app.TasksList.Items($matchedTaskIndex)}catch{}
      if($null -eq $task){throw ('Could not reacquire SigmaNEST task index '+$matchedTaskIndex+' for part "'+$targetName+'".')}
      try{$taskPart=$task.PartsList.Items($matchedPartIndex)}catch{}
      if($null -eq $taskPart){throw ('Could not reacquire SigmaNEST task-part index '+$matchedPartIndex+' for part "'+$targetName+'".')}

      $setMaterial=$false
      foreach($propertyName in @('Material','MaterialName','Mat')){
        try{
          $taskPart.$propertyName=$material
          $readBack=[string]$taskPart.$propertyName
          if($readBack.Trim().Equals($material.Trim(),[StringComparison]::OrdinalIgnoreCase)){
            $setMaterial=$true
            break
          }
        }catch{
          if(SN-IsDisconnected $_){
            throw ('SigmaNEST COM task-part disconnected while setting material property "'+$propertyName+'" for CL part "'+$targetName+'". HRESULT 0x80010108 (RPC_E_DISCONNECTED).')
          }
        }
      }

      if(-not $setMaterial){
        throw ('SigmaNEST task-part for "'+$targetName+'" does not expose a writable material property. Tried: Material, MaterialName, Mat.')
      }
    }

    if($null -ne $thickness){
      $task=$null;$taskPart=$null
      try{$task=$app.TasksList.Items($matchedTaskIndex)}catch{}
      if($null -eq $task){throw ('Could not reacquire SigmaNEST task index '+$matchedTaskIndex+' for part "'+$targetName+'" before thickness update.')}
      try{$taskPart=$task.PartsList.Items($matchedPartIndex)}catch{}
      if($null -eq $taskPart){throw ('Could not reacquire SigmaNEST task-part index '+$matchedPartIndex+' for part "'+$targetName+'" before thickness update.')}

      $setThickness=$false
      foreach($propertyName in @('Thickness','SheetThickness','Thk','MaterialThickness')){
        try{
          $taskPart.$propertyName=$thickness
          $readBack=[double]$taskPart.$propertyName
          if($readBack -eq $thickness){
            $setThickness=$true
            break
          }
        }catch{
          if(SN-IsDisconnected $_){
            throw ('SigmaNEST COM task-part disconnected while setting thickness property "'+$propertyName+'" for CL part "'+$targetName+'". HRESULT 0x80010108 (RPC_E_DISCONNECTED).')
          }
        }
      }

      if(-not $setThickness){
        throw ('SigmaNEST task-part for "'+$targetName+'" does not expose a writable thickness property. Tried: Thickness, SheetThickness, Thk, MaterialThickness.')
      }
    }
  }
}

function SN-Set-TaskPartQuantity($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){
    throw 'SigmaNEST created no TasksList entries after CreateTasksListForNewPartsInWS; cannot apply CL quantities safely.'
  }

  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $quantity=SN-Scalar-Int -value $rp.qty -default 1
    if($quantity -lt 1){$quantity=1}
    $matched=$false

    for($ti=0;$ti -lt $taskCount -and -not $matched;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}

      $taskPartCount=0
      try{$taskPartCount=[int]$task.PartsList.Count}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount -and -not $matched;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}

        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not (SN-TaskPart-MatchesRequest -taskPart $taskPart -rp $rp)){continue}

        $set=$null
        $readBack=$null
        foreach($propertyName in @('Quantity','Qty','QtyRequired','QtyReq','BatchQty')){
          try{
            $taskPart.$propertyName=$quantity
            $readBack=[double]$taskPart.$propertyName
            if($readBack -eq $quantity){
              $set=$propertyName
              break
            }
          }catch{}
        }

        if(-not $set){
          throw ('SigmaNEST task part "'+$taskPartName+'" does not expose a writable quantity property. Task index='+$ti+'.')
        }

        $matched=$true
      }
    }

    if(-not $matched){
      throw ('Could not find CL part "'+$targetName+'" inside any SigmaNEST task after task creation.')
    }
  }
}
function SN-Resolve-WS-Path($Request){
  $wsDir=[string]$Request.wsDirectory
  if([string]::IsNullOrWhiteSpace($wsDir)){
    $wsRoot=[string]$Request.wsRoot
    if([string]::IsNullOrWhiteSpace($wsRoot)){
      try{
        $parent=[IO.Directory]::GetParent([string]$Request.libraryRoot)
        if($parent){$wsRoot=$parent.FullName}
      }catch{}
    }
    if(-not [string]::IsNullOrWhiteSpace($wsRoot)){
      try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}
    }
  }
  if([string]::IsNullOrWhiteSpace($wsDir)){throw 'SigmaNEST WS output folder could not be determined.'}
  $wsDir=[IO.Path]::GetFullPath($wsDir.Trim())
  if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
  return $wsDir
}

function SN-Save-WorkspaceVerified($app,[string]$wsPath,[string]$label){
  $errors=@()
  for($attempt=1;$attempt -le 3;$attempt++){
    try{
      [void]$app.SaveWorkSpaceFile([string]$wsPath)
      Start-Sleep -Milliseconds 350
      if(Test-Path -LiteralPath $wsPath){
        return [pscustomobject]@{ok=$true;attempt=$attempt;path=$wsPath;label=$label}
      }
    }catch{
      $errors+=('Attempt '+$attempt+': '+(SN-ErrorText $_))
    }
    if($attempt -lt 3){Start-Sleep -Milliseconds 500}
  }
  throw ('Automatic '+$label+' workspace save failed: '+$wsPath+'. '+($errors -join ' | '))
}

function Invoke-SigmaNestImportGeometry($Request){
  $phase='START';$app=$null;$created=@();$partUpdates=@()
  try{
    if([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA){throw 'SigmaNEST COM requires STA.'}
    $app=New-Object -ComObject SigmaNEST.SNApp
    if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $wsDir=SN-Resolve-WS-Path -Request $Request
    $job=[string]$Request.jobName
    if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
    $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim()
    if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
    $wsPath=Join-Path $wsDir ($safe+'.ws')
    if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
    try{$app.PartsLibrary.Directory=[string]$Request.libraryRoot}catch{}
    $phase='IMPORT_PARTS'
    $before=SN-Parts-Count $app
    $queued=@()
    $queuedIndex=0
    foreach($x in @($Request.parts)){
      $source=[string]$x.sourcePath
      if([string]::IsNullOrWhiteSpace($source)){$source=[string]$x.prsPath}
      if([string]::IsNullOrWhiteSpace($source)){continue}
      $clData=[pscustomobject]@{
        sigmaMaterial=[string]$x.sigmaMaterial
        thicknessMm=$x.thicknessMm
        qty=$x.qty
      }
      $load=SN-Queue-Geometry -app $app -sourcePath $source -sourceType ([string]$x.sourceType) -clData $clData
      $qty=SN-Scalar-Int -value $x.qty -default 1
      if($qty -lt 1){$qty=1}
      $workspaceIndex=$before+$queuedIndex
      $created+=[pscustomobject]@{
        part=[string]$x.part;qty=$qty;material=[string]$x.sigmaMaterial
        thickness=SN-Scalar-Number -value $x.thicknessMm -default ([double]::NaN)
        sourcePath=$source;sourceType=[string]$x.sourceType
        matchType=[string]$x.matchType;batchMultiplier=SN-Scalar-Int -value $x.batchMultiplier -default 1
        workspaceIndex=$workspaceIndex
      }
      $queued += [pscustomobject]@{
        part=[string]$x.part;qty=$qty;sigmaMaterial=[string]$x.sigmaMaterial
        thicknessMm=$x.thicknessMm;sourcePath=$source;sourceType=[string]$x.sourceType
      }
      $queuedIndex++
    }
    if($queued.Count -eq 0){throw 'No geometry was found to import.'}
    $phase='COMMIT_IMPORTED_PARTS'
    $app.CreatePartsListForNewPartsInWS()
    $after=SN-Parts-Count $app
    if($after -lt ($before+$queued.Count)){throw ('SigmaNEST committed '+($after-$before)+' part(s) but '+$queued.Count+' were requested.')}
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts -jobName $job -linkFile ([string]$Request.clLinkFile)
    $phase='SAVE_GEOMETRY'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'geometry'
    $phase='VERIFY_SAVED_CL_DATA'
    $app.LoadWorkSpaceFile([string]$wsPath)
    $verify=SN-Verify-WorkspaceCLData -app $app -requestParts $queued
    return [pscustomobject]@{
      ok=$true;phase='IMPORT_COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count
      partUpdates=@($partUpdates);clLinkFile=[string]$Request.clLinkFile;tasksCreated=0;message=('Geometry imported, CL data applied, and saved to '+$wsPath)
      sourceCount=$created.Count
    }
  }catch{
    $exists=$false;try{$exists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false;phase=$phase;wsPath=$(if($exists){[string]$wsPath}else{''})
      parts=$created;partCount=@($created).Count;partUpdates=@($partUpdates)
      checkpointSaved=$exists;error=$_.Exception.Message
      message=$(if($exists){'Geometry workspace saved before failure: '+$wsPath}else{'Geometry import failed before a verified workspace save.'})
    }
  }finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}

function SN-Verify-TaskCLData($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){throw 'No SigmaNEST tasks are available for CL-data verification.'}

  foreach($rp in @($requestParts)){
    $foundTaskPart=$null
    for($ti=0;$ti -lt $taskCount -and $null -eq $foundTaskPart;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}
      $pc=0;try{$pc=[int]$task.PartsList.Count}catch{}
      for($pi=0;$pi -lt $pc;$pi++){
        $tp=$null;try{$tp=$task.PartsList.Items($pi)}catch{continue}
        if($null -ne $tp -and (SN-TaskPart-MatchesRequest -taskPart $tp -rp $rp)){
          $foundTaskPart=$tp
          break
        }
      }
    }
    if($null -eq $foundTaskPart){
      throw ('Post-save task verification could not map CL part "'+[string]$rp.part+'".')
    }

    $material=[string]$rp.sigmaMaterial
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $m=SN-Read-PartField -partObj $foundTaskPart -aliases @('Material','MaterialName','PartMaterial','MaterialDescription','MaterialType','Mat','MatName','SheetMaterial','StockMaterial','NestMaterial') -expectedText $material
      if($null -eq $m){throw ('Post-save task verification failed for "'+[string]$rp.part+'": material is not "'+$material+'".')}
    }

    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    if(-not [double]::IsNaN($thickness)){
      $t=SN-Read-PartField -partObj $foundTaskPart -aliases @('Thickness','SheetThickness','MaterialThickness','ThicknessValue','PartThickness','Thk','Thick') -expectedNumber $thickness
      if($null -eq $t){throw ('Post-save task verification failed for "'+[string]$rp.part+'": thickness is not '+$thickness+'mm.')}
    }

    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $q=SN-Read-PartField -partObj $foundTaskPart -aliases @('QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity','BatchQty','BatchQuantity','QtyToNest','QuantityToNest','NestQuantity') -expectedInt $qty
    if($null -eq $q){throw ('Post-save task verification failed for "'+[string]$rp.part+'": quantity is not '+$qty+'.')}
  }
  return $true
}

function SN-Set-TaskNameAndBatch($app,$requestParts){
  $taskCount=0;try{$taskCount=[int]$app.TasksList.Count}catch{}
  $map=@{};foreach($rp in @($requestParts)){$map[[string]$rp.part]=$rp}
  $results=@();$warnings=@()
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$app.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $names=@();$materials=@();$thks=@();$multis=@()
    $pc=0;try{$pc=[int]$task.PartsList.Count}catch{}
    for($pi=0;$pi -lt $pc;$pi++){
      try{$tp=$task.PartsList.Items($pi)}catch{continue}
      $pn='';try{$pn=[string]$tp.Name}catch{}
      $rp=SN-Find-RequestForTaskPart -taskPart $tp -requestParts $requestParts
      if($null -ne $rp){
        $names+=[string]$tp.Name
        if(-not [string]::IsNullOrWhiteSpace([string]$rp.sigmaMaterial)){$materials+=[string]$rp.sigmaMaterial}
        $th=[double](SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN))
        if(-not [double]::IsNaN($th)){$thks+=$th}
        foreach($tb in @($rp.taskBatches)){
          $multis+=SN-Scalar-Int -value $tb.batchMultiplier -default 1
        }
        if($multis.Count -eq 0){$multis+=SN-Scalar-Int -value $rp.batchMultiplier -default 1}
      }
    }
    if($names.Count -eq 0){continue}
    $mat=if($materials.Count){$materials[0]}else{'UNKNOWN MATERIAL'}
    $thk=if($thks.Count){$thks[0]}else{0}
    $uniqueMulti=@($multis|Sort-Object -Unique)
    $label=('{0:00} | {1} | {2:0.###}mm' -f ($ti+1),$mat,$thk)
    $labelApplied=$false;$labelProp=''
    foreach($prop in @('Name','TaskName','Description')){
      try{
        $task.$prop=$label
        $back=[string]$task.$prop
        if($back.Trim().Equals($label.Trim(),[StringComparison]::OrdinalIgnoreCase)){$labelApplied=$true;$labelProp=$prop;break}
      }catch{}
    }
    if(-not $labelApplied){$warnings+=('Task '+($ti+1)+' could not be renamed; material/thickness grouping still exists.')}
    $batchApplied=$false;$batchProp='';$batch=1
    if($uniqueMulti.Count -eq 1){
      $batch=SN-Scalar-Int -value $uniqueMulti[0] -default 1
      if($batch -lt 1){$batch=1}
      foreach($prop in @('BatchMultiplier','BatchQty','BatchQuantity','Batch')){
        try{
          $task.$prop=$batch
          $back=SN-Scalar-Int -value $task.$prop -default -1
          if($back -eq $batch){$batchApplied=$true;$batchProp=$prop;break}
        }catch{}
      }
      if(-not $batchApplied){
        foreach($pi in 0..([math]::Max(0,$pc-1))){
          try{$tp=$task.PartsList.Items($pi)}catch{continue}
          foreach($prop in @('BatchQty','BatchMultiplier','BatchQuantity','Batch')){
            try{
              $tp.$prop=$batch
              $back=SN-Scalar-Int -value $tp.$prop -default -1
              if($back -eq $batch){$batchApplied=$true;$batchProp='PART.'+$prop;break}
            }catch{}
          }
          if($batchApplied){break}
        }
      }
    }else{
      $warnings+=('Task '+($ti+1)+' has mixed CL multipliers: '+($uniqueMulti -join ', ')+'. No multiplier was guessed.')
    }
    if(-not $batchApplied -and $uniqueMulti.Count -eq 1){
      $warnings+=('Task '+($ti+1)+' could not accept Batch multiplier x'+$batch+'.')
    }
    $results+=[pscustomobject]@{
      taskIndex=$ti+1;taskName=$label;material=$mat;thickness=$thk;batchMultiplier=$batch
      batchApplied=$batchApplied;batchProperty=$batchProp;labelApplied=$labelApplied;labelProperty=$labelProp
      partCount=$names.Count;multipliers=($uniqueMulti -join ',')
    }
  }
  return [pscustomobject]@{tasks=@($results);warnings=@($warnings)}
}

function Invoke-SigmaNestAutoTask($Request){
  $phase='START';$app=$null
  try{
    if([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA){throw 'SigmaNEST COM requires STA.'}
    $wsPath=[IO.Path]::GetFullPath([string]$Request.wsPath)
    if(-not(Test-Path -LiteralPath $wsPath)){throw ('SigmaNEST WS not found: '+$wsPath)}
    $app=New-Object -ComObject SigmaNEST.SNApp
    if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}

    $phase='LOAD_WORKSPACE'
    $app.LoadWorkSpaceFile([string]$wsPath)

    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts -jobName ([IO.Path]::GetFileNameWithoutExtension($wsPath))

    $phase='AUTO_TASK'
    $app.AutoTask()
    Start-Sleep -Milliseconds 500
    $taskCount=0;try{$taskCount=[int]$app.TasksList.Count}catch{}
    if($taskCount -le 0){throw 'SigmaNEST AutoTask completed but created no tasks.'}

    $phase='APPLY_TASK_CL_DATA'
    # AutoTask groups by the part data available to SigmaNEST. Reapply the
    # authoritative CL material/thickness/quantity at task-part level and verify
    # every change before anything is saved.
    SN-Set-TaskMaterialAndThickness -app $app -requestParts $Request.parts
    SN-Set-TaskPartQuantity -app $app -requestParts $Request.parts

    $phase='LABEL_AND_BATCH'
    $taskData=SN-Set-TaskNameAndBatch -app $app -requestParts $Request.parts
    $allWarnings=@($taskData.warnings)

    $phase='SAVE'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'AutoTask'
    $app.LoadWorkSpaceFile([string]$wsPath)
    SN-Verify-WorkspaceCLData -app $app -requestParts $Request.parts | Out-Null
    SN-Verify-TaskCLData -app $app -requestParts $Request.parts | Out-Null

    $ok=($allWarnings.Count -eq 0)
    return [pscustomobject]@{
      ok=$ok
      phase='AUTOTASK_COMPLETE'
      wsPath=$wsPath
      tasksCreated=$taskCount
      taskData=$taskData.tasks
      warnings=$allWarnings
      partUpdates=@($partUpdates)
      message=$(if($ok){'AutoTask created, CL-linked, labeled and batched '+$taskCount+' task(s).'}else{'AutoTask created '+$taskCount+' task(s) with non-fatal labeling warnings; see the Release Summary.'})
    }
  }catch{
    return [pscustomobject]@{
      ok=$false
      phase=$phase
      wsPath=$wsPath
      tasksCreated=0
      taskData=@()
      warnings=@($_.Exception.Message)
      message=('AutoTask failed at '+$phase+': '+$_.Exception.Message)
    }
  }finally{
    if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}
  }
}

function Invoke-SigmaNestBuild($Request){
  $phase='START';$app=$null
  try{
    $apartment=[Threading.Thread]::CurrentThread.GetApartmentState()
    if($apartment -ne [Threading.ApartmentState]::STA){throw ('Current PowerShell thread is '+$apartment+'; SigmaNEST COM requires STA.')}
    $phase='CREATE_COM';$app=New-Object -ComObject SigmaNEST.SNApp;if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $phase='RESOLVE_WS_PATH';$wsDir=[string]$Request.wsDirectory
    if([string]::IsNullOrWhiteSpace($wsDir)){
      $wsRoot=[string]$Request.wsRoot
      if([string]::IsNullOrWhiteSpace($wsRoot)){try{$parent=[IO.Directory]::GetParent([string]$Request.libraryRoot);if($parent){$wsRoot=$parent.FullName}}catch{}}
      if(-not [string]::IsNullOrWhiteSpace($wsRoot)){try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}}
    }
    if([string]::IsNullOrWhiteSpace($wsDir)){throw 'SigmaNEST WS output folder could not be determined. Configure the SigmaNEST WS folder (PathID 0) or verify the PRS library path.'}
    $wsDir=[IO.Path]::GetFullPath($wsDir.Trim());if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
    $job=[string]$Request.jobName;if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
    $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim();if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
    $wsPath=Join-Path -Path $wsDir -ChildPath ($safe+'.ws');if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
    $phase='CONFIGURE_LIBRARY';try{$app.PartsLibrary.Directory=[string]$Request.libraryRoot}catch{}
    $phase='IMPORT_PARTS';$created=@();$queued=@()
    $beforeParts=SN-Parts-Count $app
    foreach($x in @($Request.parts)){
      $sourcePath=[string]$x.sourcePath;$sourceType=[string]$x.sourceType
      if([string]::IsNullOrWhiteSpace($sourcePath)){$sourcePath=[string]$x.prsPath;if(-not $sourceType){$sourceType='PRS'}}
      if([string]::IsNullOrWhiteSpace($sourcePath)){continue}

      $load=SN-Queue-Geometry -app $app -sourcePath $sourcePath -sourceType $sourceType
      $quantity=SN-Scalar-Int -value $x.qty -default 1;if($quantity -lt 1){$quantity=1}
      $material=[string]$x.sigmaMaterial
      $thicknessText=$(if($x.thicknessMm -ne $null -and -not [double]::IsNaN([double]$x.thicknessMm)){[string]$x.thicknessMm}else{''})
      $queued += [pscustomobject]@{
        request=$x
        sourcePath=$sourcePath
        sourceType=$sourceType
        method=$load.method
      }
      $created += [pscustomobject]@{
        part=[string]$x.part
        qty=$quantity
        material=$material
        thickness=$thicknessText
        quantityProperty='TASK_PART_PENDING'
        materialProperty='TASK_PENDING'
        thicknessProperty='TASK_PENDING'
        sourcePath=$sourcePath
        sourceType=$sourceType
        sourceProperty=''
        batchMultiplier=$x.batchMultiplier
        taskBatches=@($x.taskBatches)
      }
    }

    if($queued.Count -eq 0){
      throw 'No geometry was queued for SigmaNEST import.'
    }

    $phase='COMMIT_IMPORTED_PARTS';$app.CreatePartsListForNewPartsInWS()
    $afterParts=SN-Parts-Count $app
    if($afterParts -lt ($beforeParts+$queued.Count)){
      throw ('SigmaNEST committed '+($afterParts-$beforeParts)+' part(s) but '+$queued.Count+' part(s) were requested for import.')
    }

    # Autosave the geometry workspace immediately after the imported parts are
    # committed. This guarantees a usable .ws exists even if later Task Setup
    # automation fails or disconnects a COM proxy.
    # Apply CL production data before the first workspace checkpoint. A saved
    # checkpoint must never be the first artifact containing only PRS metadata.
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts -jobName $safe -linkFile ([string]$Request.clLinkFile)

    $phase='SAVE_GEOMETRY_CHECKPOINT'
    $saved=$false
    $saveErrors=@()
    for($attempt=1;$attempt -le 3 -and -not $saved;$attempt++){
      try{
        [void]$app.SaveWorkSpaceFile([string]$wsPath)
        Start-Sleep -Milliseconds 350
        $saved=Test-Path -LiteralPath $wsPath
      }catch{
        $saveErrors+=('Attempt '+$attempt+': '+(SN-ErrorText $_))
        if($attempt -lt 3){Start-Sleep -Milliseconds 500}
      }
    }
    if(-not $saved){
      throw ('SigmaNEST imported '+($afterParts-$beforeParts)+' part(s), CL data was applied, but the automatic .ws save did not produce a file: '+$wsPath+'. '+($saveErrors -join ' | '))
    }

    $phase='CREATE_TASKS';$app.CreateTasksListForNewPartsInWS()
    # Save again after the task list exists so the workspace remains resumable
    # if a later task attribute or quantity update fails.
    $phase='SAVE_TASK_CHECKPOINT'
    $savedTask=$false
    try{
      [void]$app.SaveWorkSpaceFile([string]$wsPath)
      Start-Sleep -Milliseconds 350
      $savedTask=Test-Path -LiteralPath $wsPath
    }catch{}
    if(-not $savedTask){
      throw ('SigmaNEST task list was created, but the automatic task checkpoint could not be confirmed on disk: '+$wsPath)
    }
    $phase='APPLY_TASK_CL_DATA'
    SN-Set-TaskMaterialAndThickness -app $app -requestParts $Request.parts
    $phase='APPLY_TASK_QUANTITIES'
    SN-Set-TaskPartQuantity -app $app -requestParts $Request.parts
    $phase='SAVE_WS'
    [void](SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'final')
    $app.LoadWorkSpaceFile([string]$wsPath)
    SN-Verify-WorkspaceCLData -app $app -requestParts $Request.parts | Out-Null
    SN-Verify-TaskCLData -app $app -requestParts $Request.parts | Out-Null
    try{$app.RefreshTreeView()}catch{};try{$app.Redraw()}catch{}
    return [pscustomobject]@{ok=$true;creatorVersion='DIRECT-COM-2.13.0';phase='COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count;clLinkFile=[string]$Request.clLinkFile;message=('SigmaNEST WS created with CL-linked part data: '+$wsPath)}
  }catch{
    $checkpointExists=$false
    try{$checkpointExists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false
      creatorVersion='DIRECT-COM-2.13.0'
      phase=$phase
      error=$_.Exception.Message
      category=$_.CategoryInfo.ToString()
      wsPath=$(if($checkpointExists){[string]$wsPath}else{''})
      parts=$created
      partCount=@($created).Count
      checkpointSaved=$checkpointExists
      message=$(if($checkpointExists){'SigmaNEST created a geometry/workspace checkpoint before the failure: '+[string]$wsPath}else{'SigmaNEST build failed before a workspace checkpoint was saved.'})
    }
  }
  finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
},'') -replace '(?i)DXFfunction SN-Find-WorkspacePartExact($app,$targetName,$sourcePath='',$usedIndices=@()){
  $targetNorm=SN-Normalize-PartIdentity $targetName
  $sourceStem=''
  try{$sourceStem=SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension([string]$sourcePath))}catch{}
  $sourceStemNoPrs=SN-Normalize-PartIdentity ($sourceStem -replace '(?i)PRS$','')
  $count=SN-Parts-Count $app
  $candidates=@()
  for($i=0;$i -lt $count;$i++){
    if(@($usedIndices) -contains $i){continue}
    $part=$null
    try{$part=$app.PartsList.Items($i)}catch{continue}
    if($null -eq $part){continue}
    $id=SN-Get-PartIdentity $part
    $score=0
    if($targetNorm -and $id.normalizedName -eq $targetNorm){$score=100}
    if($sourceStemNoPrs -and $id.normalizedFile -eq $sourceStemNoPrs){$score=[math]::Max($score,95)}
    if($targetNorm -and $id.normalizedDrawing -eq $targetNorm){$score=[math]::Max($score,90)}
    if($sourceStemNoPrs -and $id.normalizedDrawing -eq $sourceStemNoPrs){$score=[math]::Max($score,85)}
    if($score -gt 0){
      $candidates += [pscustomobject]@{part=$part;index=$i;score=$score;identity=$id}
    }
  }
  if($candidates.Count -eq 0){return $null}
  $top=@($candidates|Sort-Object @{Expression={[int]$_.score};Descending=$true},@{Expression={[int]$_.index};Descending=$false})
  if($top.Count -gt 1 -and [int]$top[0].score -eq [int]$top[1].score){
    # Prefer an exact normalized Name match over filename/drawing fallbacks.
    $exact=@($top|Where-Object {$_.identity.normalizedName -eq $targetNorm})
    if($exact.Count -eq 1){return $exact[0]}
    if($exact.Count -gt 1){throw ('Multiple SigmaNEST workspace parts match CL part "'+$targetName+'".')}
  }
  return $top[0]
}
function SN-Verify-PartIdentity($partObj,$targetName,$sourcePath=''){
  $id=SN-Get-PartIdentity $partObj
  $targetNorm=SN-Normalize-PartIdentity $targetName
  $sourceNorm=''
  try{$sourceNorm=SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension([string]$sourcePath))}catch{}
  $sourceNorm=SN-Normalize-PartIdentity ($sourceNorm -replace '(?i)PRS$','')
  if($targetNorm -and $id.normalizedName -eq $targetNorm){return $true}
  if($targetNorm -and $id.normalizedDrawing -eq $targetNorm){return $true}
  if($sourceNorm -and $id.normalizedFile -eq $sourceNorm){return $true}
  if($sourceNorm -and $id.normalizedDrawing -eq $sourceNorm){return $true}
  return $false
}
function SN-Apply-WorkspacePartData($app,$requestParts){
  $updated=@()
  $usedIndices=@()
  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $sourcePath=[string]$rp.sourcePath
    $found=SN-Find-WorkspacePartExact -app $app -targetName $targetName -sourcePath $sourcePath -usedIndices $usedIndices
    if($null -eq $found){
      throw ('Could not map imported SigmaNEST geometry back to CL part "'+$targetName+'". The imported PartsList order is not trusted.')
    }
    if(-not (SN-Verify-PartIdentity -partObj $found.part -targetName $targetName -sourcePath $sourcePath)){
      throw ('SigmaNEST workspace identity verification failed for CL part "'+$targetName+'". Refusing to overwrite another part.')
    }
    $usedIndices += [int]$found.index
    $partObj=$found.part
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $material=[string]$rp.sigmaMaterial
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    $row=[ordered]@{
      part=$targetName
      index=[int]$found.index
      qty=$qty
      material=$material
      thickness=$thickness
      quantityProperty=''
      materialProperty=''
      thicknessProperty=''
      warnings=@()
    }

    if(-not [string]::IsNullOrWhiteSpace($material)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @('Material','MaterialName','Mat','MatName','MaterialType') -value $material -expectedText $material
      if(-not $setInfo){
        throw ('CL material "'+$material+'" could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.materialProperty=[string]$setInfo.path
    }

    if(-not [double]::IsNaN($thickness)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -value $thickness -expectedNumber $thickness
      if(-not $setInfo){
        throw ('CL thickness '+$thickness+'mm could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.thicknessProperty=[string]$setInfo.path
    }

    $setInfo=SN-Set-PartField -partObj $partObj -names @('QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -value $qty -expectedInt $qty
    if(-not $setInfo){
      throw ('CL quantity '+$qty+' could not be written and verified on SigmaNEST part "'+$targetName+'".')
    }
    $row.quantityProperty=[string]$setInfo.path
    $updated += [pscustomobject]$row
  }
  @($updated)
}
function SN-Read-PartField($partObj,[string[]]$aliases,[string]$expectedText='',$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $partObj){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$partObj;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 4){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    foreach($prop in @(SN-ComPropertyNames $obj)){
      if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
      try{
        $v=$obj.$prop
        if($expectedText -ne ''){
          if(([string]$v).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=[string]$v}
          }
        }elseif(-not [double]::IsNaN($expectedNumber)){
          $n=SN-Scalar-Number -value $v -default ([double]::NaN)
          if(-not [double]::IsNaN([double]$n) -and [double]$n -eq $expectedNumber){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=$n}
          }
        }elseif($expectedInt -ne -2147483648){
          $n=SN-Scalar-Int -value $v -default -2147483648
          if($n -eq $expectedInt){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=$n}
          }
        }else{
          return [pscustomobject]@{path=($node.path+'.'+$prop);value=$v}
        }
      }catch{}
    }

    if($depth -ge 4){continue}
    foreach($prop in @(SN-ComPropertyNames $obj)){
      if($prop -in @('OwnerList','ParentObject','Parent','Application','Owner','Geometry','PartPolyLinesList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Verify-WorkspaceCLData($app,$requestParts){
  foreach($rp in @($requestParts)){
    $target=[string]$rp.part
    $source=[string]$rp.sourcePath
    $found=SN-Find-WorkspacePartExact -app $app -targetName $target -sourcePath $source -usedIndices @()
    if($null -eq $found){throw ('Post-save verification could not find SigmaNEST part "'+$target+'".')}
    $material=[string]$rp.sigmaMaterial
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','Mat','MatName','MaterialType') -expectedText $material
      if($null -eq $m){throw ('Post-save verification failed for "'+$target+'": material is not "'+$material+'".')}
    }
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    if(-not [double]::IsNaN($thickness)){
      $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -expectedNumber $thickness
      if($null -eq $t){throw ('Post-save verification failed for "'+$target+'": thickness is not '+$thickness+'mm.')}
    }
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $q=SN-Read-PartField -partObj $found.part -aliases @('QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $qty
    if($null -eq $q){throw ('Post-save verification failed for "'+$target+'": quantity is not '+$qty+'.')}
  }
  return $true
}
function SN-Parts-Count($app){try{return [int]$app.PartsList.Count}catch{return 0}}

function SN-Get-NewPart($app,$beforeCount,[string]$label){
  $beforeCount=SN-Scalar-Int -value $beforeCount -default 0;
  $afterCount=SN-Parts-Count $app
  if($afterCount -le $beforeCount){throw ('SigmaNEST loaded "'+$label+'" but did not add it to the workspace PartsList.')}
  try{$last=$null;foreach($item in $app.PartsList){$last=$item};if($null -ne $last){return $last}}catch{}
  foreach($member in @('Item','Items','get_Item')){
    try{$obj=$app.PartsList.GetType().InvokeMember($member,[Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::GetProperty,$null,$app.PartsList,@([int]($afterCount-1)));if($null -ne $obj){return $obj}}catch{}
  }
  throw ('SigmaNEST added "'+$label+'" but the new PartsList item could not be accessed. PartsList.Count='+$afterCount)
}

function SN-AllMemberDefinitions($obj){
  if($null -eq $obj){return @()}
  try{
    return @($obj | Get-Member -MemberType Methods,Properties,ParameterizedProperty,CodeProperty,NoteProperty,ScriptProperty,Fields -ErrorAction Stop |
      ForEach-Object {[pscustomobject]@{Name=[string]$_.Name;Type=[string]$_.MemberType;Definition=[string]$_.Definition}})
  }catch{return @()}
}
function SN-ComPropertyNamesAll($obj){
  return @((SN-AllMemberDefinitions $obj)|Where-Object {$_.Type -match 'Property'}|Select-Object -ExpandProperty Name -Unique)
}
function SN-InvokeSingleArgMember($obj,[string]$methodName,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $obj -or [string]::IsNullOrWhiteSpace($methodName)){return $null}
  try{
    $defs=@((SN-AllMemberDefinitions $obj)|Where-Object {$_.Name -eq $methodName})
    foreach($d in $defs){
      $def=[string]$d.Definition
      if($def -notmatch '\\(([^)]*)\\)'){continue}
      $inside=$Matches[1].Trim()
      if($inside -match ','){continue}
      if($inside -and $inside -notmatch '(?i)optional|paramarray'){ }
      try{
        $result=$obj.GetType().InvokeMember($methodName,[Reflection.BindingFlags]::InvokeMethod,$null,$obj,@($value))
        if($expectedText -ne ''){
          if($result -ne $null -and ([string]$result).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $methodName}
          try{
            $back=$obj.GetType().InvokeMember($methodName,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())
            if(([string]$back).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $methodName}
          }catch{}
        }elseif(-not [double]::IsNaN($expectedNumber)){
          try{if(([double]$result) -eq $expectedNumber){return $methodName}}catch{}
        }elseif($expectedInt -ne -2147483648){
          try{if((SN-Scalar-Int $result -default -2147483648) -eq $expectedInt){return $methodName}}catch{}
        }else{return $methodName}
      }catch{}
    }
  }catch{}
  return $null
}
function SN-Try-SetImportSetting($settings,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $settings){return $null}
  # 1) Standard writable/parameterized properties.
  foreach($prop in @(SN-ComPropertyNamesAll $settings)){
    if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
    try{
      $settings.$prop=$value
      if($expectedText -ne '' -and ([string]$settings.$prop).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      if(-not [double]::IsNaN($expectedNumber) -and (SN-Scalar-Number $settings.$prop -default ([double]::NaN)) -eq $expectedNumber){return $prop}
      if($expectedInt -ne -2147483648 -and (SN-Scalar-Int $settings.$prop -default -2147483648) -eq $expectedInt){return $prop}
    }catch{}
    try{
      $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::SetProperty,$null,$settings,@($value))|Out-Null
      if($expectedText -ne ''){
        $back=[string]$settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@())
        if($back.Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        $back=SN-Scalar-Number $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@()) -default ([double]::NaN)
        if(-not [double]::IsNaN($back) -and $back -eq $expectedNumber){return $prop}
      }elseif($expectedInt -ne -2147483648){
        $back=SN-Scalar-Int $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@()) -default -2147483648
        if($back -eq $expectedInt){return $prop}
      }
    }catch{}
  }

  # 2) Setter-style methods such as SetMaterial / SetThickness / SetQuantity.
  foreach($method in @(SN-AllMemberDefinitions $settings | Where-Object {$_.Type -eq 'Method'} | Select-Object -ExpandProperty Name -Unique)){
    $norm=SN-NormalizeFieldName $method
    foreach($alias in @($aliases)){
      $an=SN-NormalizeFieldName $alias
      if($norm -eq $an -or $norm -eq ('SET'+$an) -or $norm -eq ('SET'+$an+'VALUE') -or $norm -eq ('SET'+$an+'NAME') -or ($norm.StartsWith('SET') -and $norm.Contains($an))){
        $hit=SN-InvokeSingleArgMember -obj $settings -methodName $method -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
        if($hit){return ('METHOD.'+$hit)}
      }
    }
  }

  # 3) Generic field/value methods used by COM import-mapping objects.
  foreach($method in @('SetValue','SetField','SetParameter','SetProperty','SetImportValue','SetOption')){
    $defs=@((SN-AllMemberDefinitions $settings)|Where-Object {$_.Name -eq $method -and $_.Type -eq 'Method'})
    foreach($d in $defs){
      try{
        if(([string]$d.Definition) -match '\\([^)]*\\)' -and [string]$d.Definition -notmatch ','){
          $r=$settings.GetType().InvokeMember($method,[Reflection.BindingFlags]::InvokeMethod,$null,$settings,@($value))
          if($expectedText -ne '' -and $r -ne $null -and ([string]$r).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return ('GENERIC.'+$method)}
          if($expectedInt -ne -2147483648 -and (SN-Scalar-Int $r -default -2147483648) -eq $expectedInt){return ('GENERIC.'+$method)}
          if(-not [double]::IsNaN($expectedNumber) -and (SN-Scalar-Number $r -default ([double]::NaN)) -eq $expectedNumber){return ('GENERIC.'+$method)}
        }
      }catch{}
    }
  }
  return $null
}
function SN-Get-ImportSettings($adapter){
  if($null -ne $adapter){
    foreach($prop in @('PartImportSettings','ImportSettings','Settings')){
      try{
        $v=$adapter.$prop
        if($null -ne $v){return [pscustomobject]@{settings=$v;property=$prop}}
      }catch{}
    }
  }
  foreach($progId in @('SigmaNEST.SNPartImportSettings','SigmaNEST.SNPartImportSetting')){
    try{
      $v=New-Object -ComObject $progId -ErrorAction Stop
      if($null -ne $v){return [pscustomobject]@{settings=$v;property=$progId}}
    }catch{}
  }
  return $null
}
function SN-Get-ImportDiagnostics($adapter,$settings){
  $parts=@()
  if($null -ne $adapter){
    $parts+='Adapter methods/properties:'
    foreach($d in @(SN-AllMemberDefinitions $adapter)){$parts+=('  '+$d.Name+' :: '+$d.Definition)}
  }
  if($null -ne $settings){
    $parts+='PartImportSettings methods/properties:'
    foreach($d in @(SN-AllMemberDefinitions $settings)){$parts+=('  '+$d.Name+' :: '+$d.Definition)}
    foreach($prop in @(SN-ComPropertyNamesAll $settings)){
      try{
        $child=$settings.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array] -and $child.Count -gt 1){continue}
        $parts+=('  Nested '+$prop+':')
        foreach($d in @(SN-AllMemberDefinitions $child)){$parts+=('    '+$d.Name+' :: '+$d.Definition)}
      }catch{}
    }
  }
  return @($parts)
}
function SN-SetImportSettingDeep($root,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $root){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$root;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 5){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    $hit=SN-Try-SetImportSetting -settings $obj -aliases $aliases -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
    if($hit){return [pscustomobject]@{path=($node.path+'.'+$hit)}}

    if($depth -ge 5){continue}
    foreach($prop in @(SN-ComPropertyNamesAll $obj)){
      if($prop -in @('Owner','Parent','ParentObject','Application','PartImportSettings','SolidsList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Configure-PRS-ImportSettings($settings,$clData){
  $material=[string]$clData.sigmaMaterial
  $thickness=SN-Scalar-Number -value $clData.thicknessMm -default ([double]::NaN)
  $qty=SN-Scalar-Int -value $clData.qty -default 1
  if($qty -lt 1){$qty=1}
  $result=[ordered]@{material='';thickness='';quantity='';warnings=@()}
  if(-not [string]::IsNullOrWhiteSpace($material)){
    $r=SN-SetImportSettingDeep -root $settings -aliases @('Material','MaterialName','PartMaterial','MaterialType','MaterialString') -value $material -expectedText $material
    if($r){$result.material=$r.path}
  }
  if(-not [double]::IsNaN($thickness)){
    $r=SN-SetImportSettingDeep -root $settings -aliases @('Thickness','SheetThickness','MaterialThickness','MaterialThk','Thk','Thick') -value $thickness -expectedNumber $thickness
    if($r){$result.thickness=$r.path}
  }
  $r=SN-SetImportSettingDeep -root $settings -aliases @('QtyOrdered','Quantity','PartQuantity','NumberToNest','NumberToLoad','QuantityToNest','QtyToNest','BatchQuantity') -value $qty -expectedInt $qty
  if($r){$result.quantity=$r.path}
  if([string]::IsNullOrWhiteSpace([string]$result.material)){ $result.warnings+='Import settings did not expose a verified material field.' }
  if(-not [double]::IsNaN($thickness) -and [string]::IsNullOrWhiteSpace([string]$result.thickness)){ $result.warnings+='Import settings did not expose a verified thickness field.' }
  if([string]::IsNullOrWhiteSpace([string]$result.quantity)){ $result.warnings+='Import settings did not expose a verified quantity field.' }
  return [pscustomobject]$result
}
function SN-Try-ImportPRS-WithSettings($app,[string]$sourcePath,$clData){
  $before=SN-Parts-Count $app
  $adapter=$null
  $info=$null
  $diagnostics=@()
  try{$adapter=New-Object -ComObject SigmaNEST.SNPartExportImport -ErrorAction Stop}catch{
    $diagnostics+='SNPartExportImport unavailable: '+(SN-ErrorText $_)
  }
  if($null -ne $adapter){
    $info=SN-Get-ImportSettings -adapter $adapter
    if($info){
      $cfg=SN-Configure-PRS-ImportSettings -settings $info.settings -clData $clData
      if($cfg.warnings.Count -eq 0){
        foreach($methodName in @('ImportPartWithFeedback','ImportPart2','ImportPart','ImportAsParts')){
          $members=@(SN-ComMethodNames $adapter)
          if($members -notcontains $methodName){continue}
          foreach($args in @(
            @([string]$sourcePath,$info.settings),
            @([string]$sourcePath,$info.settings,$false),
            @([string]$sourcePath)
          )){
            try{
              $result=SN-Invoke-ComMethod -obj $adapter -name $methodName -args $args
              Start-Sleep -Milliseconds 150
              $after=SN-Parts-Count $app
              if($after -gt $before){
                return [pscustomobject]@{ok=$true;method=('SNPartExportImport.'+$methodName);settings=$cfg}
              }
              if($result -is [System.Object[]]){
                $result=@($result|Where-Object {$null -ne $_})[0]
              }
              if($result -and $result.PSObject.Properties.Name -contains 'Part'){
                try{
                  $part=$result.Part
                  if($null -ne $part){return [pscustomobject]@{ok=$true;method=('SNPartExportImport.'+$methodName);settings=$cfg;pendingPart=$part}}
                }catch{}
              }
            }catch{
              $diagnostics+=($methodName+' '+$args.Count+' args: '+(SN-ErrorText $_))
            }
          }
        }
      }else{
        $diagnostics+=($cfg.warnings -join ' ')
      }
    }else{
      $diagnostics+='SNPartExportImport has no accessible PartImportSettings/ImportSettings/Settings object.'
    }
  }
  if($null -ne $adapter){
    try{$diagnostics+='Adapter methods: '+((SN-ComMethodNames $adapter) -join ', ')}catch{}
    try{$diagnostics+='Adapter properties: '+((SN-ComPropertyNames $adapter) -join ', ')}catch{}
  }
  try{if($adapter){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($adapter)}}catch{}
  if($null -ne $adapter -and $null -ne $info -and $null -ne $info.settings){
    $diagnostics += @(SN-Get-ImportDiagnostics -adapter $adapter -settings $info.settings)
  }
  return [pscustomobject]@{ok=$false;diagnostics=$diagnostics}
}
function SN-Queue-Geometry($app,[string]$sourcePath,[string]$sourceType,$clData=$null){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $errors=@()

  if($sourceType -eq 'PRS' -and $null -ne $clData){
    $prsTry=SN-Try-ImportPRS-WithSettings -app $app -sourcePath $sourcePath -clData $clData
    if($prsTry.ok){
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method=$prsTry.method;preImportSettings=$prsTry.settings}
    }
    if($prsTry.diagnostics.Count){$errors+=($prsTry.diagnostics -join ' | ')}
  }

  try{
    # LoadPart is retained as a geometry-loader fallback for DXF and for
    # installations where PRS import is known to work. PRS is NOT allowed
    # to fall through silently when CL import-time settings were unavailable.
    if($sourceType -ne 'PRS'){
      [void]$app.LoadPart([string]$sourcePath)
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method='LoadPart'}
    }
  }catch{
    $errors+=('LoadPart: '+(SN-ErrorText $_))
  }

  if($sourceType -eq 'DXF'){
    try{
      $settings=$null;try{$settings=New-Object -ComObject SigmaNEST.SNPartImportSettings}catch{}
      foreach($importId in @(0,1,2,3)){
        try{
          $result=SN-Invoke-ComMethod $app.PartsList 'Import' @([string]$sourcePath,[int]$importId,$settings)
          return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method=('PartsList.Import '+$importId)}
        }catch{
          $errors+=('PartsList.Import id '+$importId+': '+(SN-ErrorText $_))
        }
      }
    }catch{
      $errors+=('DXF import fallback: '+(SN-ErrorText $_))
    }
  }

  if($sourceType -eq 'PRS'){
    throw ('SigmaNEST could not perform a verified CL-data PRS import for "'+$label+'". The PRS was NOT imported because its original material/thickness/quantity cannot be safely accepted as production data. '+($errors -join ' | '))
  }
  throw ('SigmaNEST could not load "'+$label+'". '+($errors -join ' | '))
}


function SN-Set-TaskMaterialAndThickness($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){
    throw 'SigmaNEST created no TasksList entries after CreateTasksListForNewPartsInWS; cannot apply CL material/thickness safely.'
  }

  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $material=[string]$rp.sigmaMaterial
    $thickness=$null
    try{
      if($rp.thicknessMm -ne $null -and -not [double]::IsNaN([double]$rp.thicknessMm)){
        $thickness=[double]$rp.thicknessMm
      }
    }catch{}

    $matched=$false
    $matchedTaskIndex=-1
    $matchedPartIndex=-1

    for($ti=0;$ti -lt $taskCount -and -not $matched;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}

      $taskPartCount=0
      try{$taskPartCount=[int]$task.PartsList.Count}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}
        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not $taskPartName.Equals($targetName,[StringComparison]::OrdinalIgnoreCase)){continue}

        $matched=$true
        $matchedTaskIndex=$ti
        $matchedPartIndex=$pi
        break
      }
    }

    if(-not $matched){
      throw ('Could not find CL part "'+$targetName+'" inside any SigmaNEST task after task creation.')
    }

    # Reacquire the COM task/part immediately before each mutation. SigmaNEST
    # may rebuild the task tree when Task Setup data changes, invalidating
    # previously returned COM interfaces.
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $task=$null;$taskPart=$null
      try{$task=$app.TasksList.Items($matchedTaskIndex)}catch{}
      if($null -eq $task){throw ('Could not reacquire SigmaNEST task index '+$matchedTaskIndex+' for part "'+$targetName+'".')}
      try{$taskPart=$task.PartsList.Items($matchedPartIndex)}catch{}
      if($null -eq $taskPart){throw ('Could not reacquire SigmaNEST task-part index '+$matchedPartIndex+' for part "'+$targetName+'".')}

      $setMaterial=$false
      foreach($propertyName in @('Material','MaterialName','Mat')){
        try{
          $taskPart.$propertyName=$material
          $readBack=[string]$taskPart.$propertyName
          if($readBack.Trim().Equals($material.Trim(),[StringComparison]::OrdinalIgnoreCase)){
            $setMaterial=$true
            break
          }
        }catch{
          if(SN-IsDisconnected $_){
            throw ('SigmaNEST COM task-part disconnected while setting material property "'+$propertyName+'" for CL part "'+$targetName+'". HRESULT 0x80010108 (RPC_E_DISCONNECTED).')
          }
        }
      }

      if(-not $setMaterial){
        throw ('SigmaNEST task-part for "'+$targetName+'" does not expose a writable material property. Tried: Material, MaterialName, Mat.')
      }
    }

    if($null -ne $thickness){
      $task=$null;$taskPart=$null
      try{$task=$app.TasksList.Items($matchedTaskIndex)}catch{}
      if($null -eq $task){throw ('Could not reacquire SigmaNEST task index '+$matchedTaskIndex+' for part "'+$targetName+'" before thickness update.')}
      try{$taskPart=$task.PartsList.Items($matchedPartIndex)}catch{}
      if($null -eq $taskPart){throw ('Could not reacquire SigmaNEST task-part index '+$matchedPartIndex+' for part "'+$targetName+'" before thickness update.')}

      $setThickness=$false
      foreach($propertyName in @('Thickness','SheetThickness','Thk','MaterialThickness')){
        try{
          $taskPart.$propertyName=$thickness
          $readBack=[double]$taskPart.$propertyName
          if($readBack -eq $thickness){
            $setThickness=$true
            break
          }
        }catch{
          if(SN-IsDisconnected $_){
            throw ('SigmaNEST COM task-part disconnected while setting thickness property "'+$propertyName+'" for CL part "'+$targetName+'". HRESULT 0x80010108 (RPC_E_DISCONNECTED).')
          }
        }
      }

      if(-not $setThickness){
        throw ('SigmaNEST task-part for "'+$targetName+'" does not expose a writable thickness property. Tried: Thickness, SheetThickness, Thk, MaterialThickness.')
      }
    }
  }
}

function SN-Set-TaskPartQuantity($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){
    throw 'SigmaNEST created no TasksList entries after CreateTasksListForNewPartsInWS; cannot apply CL quantities safely.'
  }

  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $quantity=SN-Scalar-Int -value $rp.qty -default 1
    if($quantity -lt 1){$quantity=1}
    $matched=$false

    for($ti=0;$ti -lt $taskCount -and -not $matched;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}

      $taskPartCount=0
      try{$taskPartCount=[int]$task.PartsList.Count}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount -and -not $matched;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}

        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not $taskPartName.Equals($targetName,[StringComparison]::OrdinalIgnoreCase)){continue}

        $set=$null
        $readBack=$null
        foreach($propertyName in @('Quantity','Qty','QtyRequired','QtyReq','BatchQty')){
          try{
            $taskPart.$propertyName=$quantity
            $readBack=[double]$taskPart.$propertyName
            if($readBack -eq $quantity){
              $set=$propertyName
              break
            }
          }catch{}
        }

        if(-not $set){
          throw ('SigmaNEST task part "'+$taskPartName+'" does not expose a writable quantity property. Task index='+$ti+'.')
        }

        $matched=$true
      }
    }

    if(-not $matched){
      throw ('Could not find CL part "'+$targetName+'" inside any SigmaNEST task after task creation.')
    }
  }
}
function SN-Resolve-WS-Path($Request){
  $wsDir=[string]$Request.wsDirectory
  if([string]::IsNullOrWhiteSpace($wsDir)){
    $wsRoot=[string]$Request.wsRoot
    if([string]::IsNullOrWhiteSpace($wsRoot)){
      try{
        $parent=[IO.Directory]::GetParent([string]$Request.libraryRoot)
        if($parent){$wsRoot=$parent.FullName}
      }catch{}
    }
    if(-not [string]::IsNullOrWhiteSpace($wsRoot)){
      try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}
    }
  }
  if([string]::IsNullOrWhiteSpace($wsDir)){throw 'SigmaNEST WS output folder could not be determined.'}
  $wsDir=[IO.Path]::GetFullPath($wsDir.Trim())
  if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
  return $wsDir
}

function SN-Save-WorkspaceVerified($app,[string]$wsPath,[string]$label){
  $errors=@()
  for($attempt=1;$attempt -le 3;$attempt++){
    try{
      [void]$app.SaveWorkSpaceFile([string]$wsPath)
      Start-Sleep -Milliseconds 350
      if(Test-Path -LiteralPath $wsPath){
        return [pscustomobject]@{ok=$true;attempt=$attempt;path=$wsPath;label=$label}
      }
    }catch{
      $errors+=('Attempt '+$attempt+': '+(SN-ErrorText $_))
    }
    if($attempt -lt 3){Start-Sleep -Milliseconds 500}
  }
  throw ('Automatic '+$label+' workspace save failed: '+$wsPath+'. '+($errors -join ' | '))
}

function Invoke-SigmaNestImportGeometry($Request){
  $phase='START';$app=$null;$created=@();$partUpdates=@()
  try{
    if([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA){throw 'SigmaNEST COM requires STA.'}
    $app=New-Object -ComObject SigmaNEST.SNApp
    if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $wsDir=SN-Resolve-WS-Path -Request $Request
    $job=[string]$Request.jobName
    if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
    $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim()
    if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
    $wsPath=Join-Path $wsDir ($safe+'.ws')
    if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
    try{$app.PartsLibrary.Directory=[string]$Request.libraryRoot}catch{}
    $phase='IMPORT_PARTS'
    $before=SN-Parts-Count $app
    $queued=@()
    $queuedIndex=0
    foreach($x in @($Request.parts)){
      $source=[string]$x.sourcePath
      if([string]::IsNullOrWhiteSpace($source)){$source=[string]$x.prsPath}
      if([string]::IsNullOrWhiteSpace($source)){continue}
      $clData=[pscustomobject]@{
        sigmaMaterial=[string]$x.sigmaMaterial
        thicknessMm=$x.thicknessMm
        qty=$x.qty
      }
      $load=SN-Queue-Geometry -app $app -sourcePath $source -sourceType ([string]$x.sourceType) -clData $clData
      $qty=SN-Scalar-Int -value $x.qty -default 1
      if($qty -lt 1){$qty=1}
      $workspaceIndex=$before+$queuedIndex
      $created+=[pscustomobject]@{
        part=[string]$x.part;qty=$qty;material=[string]$x.sigmaMaterial
        thickness=SN-Scalar-Number -value $x.thicknessMm -default ([double]::NaN)
        sourcePath=$source;sourceType=[string]$x.sourceType
        matchType=[string]$x.matchType;batchMultiplier=SN-Scalar-Int -value $x.batchMultiplier -default 1
        workspaceIndex=$workspaceIndex
      }
      $queued += [pscustomobject]@{
        part=[string]$x.part;qty=$qty;sigmaMaterial=[string]$x.sigmaMaterial
        thicknessMm=$x.thicknessMm;sourcePath=$source;sourceType=[string]$x.sourceType
      }
      $queuedIndex++
    }
    if($queued.Count -eq 0){throw 'No geometry was found to import.'}
    $phase='COMMIT_IMPORTED_PARTS'
    $app.CreatePartsListForNewPartsInWS()
    $after=SN-Parts-Count $app
    if($after -lt ($before+$queued.Count)){throw ('SigmaNEST committed '+($after-$before)+' part(s) but '+$queued.Count+' were requested.')}
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts
    $phase='SAVE_GEOMETRY'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'geometry'
    $phase='VERIFY_SAVED_CL_DATA'
    $app.LoadWorkSpaceFile([string]$wsPath)
    $verify=SN-Verify-WorkspaceCLData -app $app -requestParts $queued
    return [pscustomobject]@{
      ok=$true;phase='IMPORT_COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count
      partUpdates=@($partUpdates);tasksCreated=0;message=('Geometry imported and saved to '+$wsPath)
      sourceCount=$created.Count
    }
  }catch{
    $exists=$false;try{$exists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false;phase=$phase;wsPath=$(if($exists){[string]$wsPath}else{''})
      parts=$created;partCount=@($created).Count;partUpdates=@($partUpdates)
      checkpointSaved=$exists;error=$_.Exception.Message
      message=$(if($exists){'Geometry workspace saved before failure: '+$wsPath}else{'Geometry import failed before a verified workspace save.'})
    }
  }finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}

function SN-Set-TaskNameAndBatch($app,$requestParts){
  $taskCount=0;try{$taskCount=[int]$app.TasksList.Count}catch{}
  $map=@{};foreach($rp in @($requestParts)){$map[[string]$rp.part]=$rp}
  $results=@();$warnings=@()
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$app.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $names=@();$materials=@();$thks=@();$multis=@()
    $pc=0;try{$pc=[int]$task.PartsList.Count}catch{}
    for($pi=0;$pi -lt $pc;$pi++){
      try{$tp=$task.PartsList.Items($pi)}catch{continue}
      $pn='';try{$pn=[string]$tp.Name}catch{}
      if($map.ContainsKey($pn)){
        $rp=$map[$pn];$names+=$pn
        if(-not [string]::IsNullOrWhiteSpace([string]$rp.sigmaMaterial)){$materials+=[string]$rp.sigmaMaterial}
        $th=[double](SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN))
        if(-not [double]::IsNaN($th)){$thks+=$th}
        foreach($tb in @($rp.taskBatches)){
          $multis+=SN-Scalar-Int -value $tb.batchMultiplier -default 1
        }
        if($multis.Count -eq 0){$multis+=SN-Scalar-Int -value $rp.batchMultiplier -default 1}
      }
    }
    if($names.Count -eq 0){continue}
    $mat=if($materials.Count){$materials[0]}else{'UNKNOWN MATERIAL'}
    $thk=if($thks.Count){$thks[0]}else{0}
    $uniqueMulti=@($multis|Sort-Object -Unique)
    $label=('{0:00} | {1} | {2:0.###}mm' -f ($ti+1),$mat,$thk)
    $labelApplied=$false;$labelProp=''
    foreach($prop in @('Name','TaskName','Description')){
      try{
        $task.$prop=$label
        $back=[string]$task.$prop
        if($back.Trim().Equals($label.Trim(),[StringComparison]::OrdinalIgnoreCase)){$labelApplied=$true;$labelProp=$prop;break}
      }catch{}
    }
    if(-not $labelApplied){$warnings+=('Task '+($ti+1)+' could not be renamed; material/thickness grouping still exists.')}
    $batchApplied=$false;$batchProp='';$batch=1
    if($uniqueMulti.Count -eq 1){
      $batch=SN-Scalar-Int -value $uniqueMulti[0] -default 1
      if($batch -lt 1){$batch=1}
      foreach($prop in @('BatchMultiplier','BatchQty','BatchQuantity','Batch')){
        try{
          $task.$prop=$batch
          $back=SN-Scalar-Int -value $task.$prop -default -1
          if($back -eq $batch){$batchApplied=$true;$batchProp=$prop;break}
        }catch{}
      }
      if(-not $batchApplied){
        foreach($pi in 0..([math]::Max(0,$pc-1))){
          try{$tp=$task.PartsList.Items($pi)}catch{continue}
          foreach($prop in @('BatchQty','BatchMultiplier','BatchQuantity','Batch')){
            try{
              $tp.$prop=$batch
              $back=SN-Scalar-Int -value $tp.$prop -default -1
              if($back -eq $batch){$batchApplied=$true;$batchProp='PART.'+$prop;break}
            }catch{}
          }
          if($batchApplied){break}
        }
      }
    }else{
      $warnings+=('Task '+($ti+1)+' has mixed CL multipliers: '+($uniqueMulti -join ', ')+'. No multiplier was guessed.')
    }
    if(-not $batchApplied -and $uniqueMulti.Count -eq 1){
      $warnings+=('Task '+($ti+1)+' could not accept Batch multiplier x'+$batch+'.')
    }
    $results+=[pscustomobject]@{
      taskIndex=$ti+1;taskName=$label;material=$mat;thickness=$thk;batchMultiplier=$batch
      batchApplied=$batchApplied;batchProperty=$batchProp;labelApplied=$labelApplied;labelProperty=$labelProp
      partCount=$names.Count;multipliers=($uniqueMulti -join ',')
    }
  }
  return [pscustomobject]@{tasks=@($results);warnings=@($warnings)}
}

function Invoke-SigmaNestAutoTask($Request){
  $phase='START';$app=$null
  try{
    if([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA){throw 'SigmaNEST COM requires STA.'}
    $wsPath=[IO.Path]::GetFullPath([string]$Request.wsPath)
    if(-not(Test-Path -LiteralPath $wsPath)){throw ('SigmaNEST WS not found: '+$wsPath)}
    $app=New-Object -ComObject SigmaNEST.SNApp
    if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $phase='LOAD_WORKSPACE'
    $app.LoadWorkSpaceFile([string]$wsPath)
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $queued
    $phase='AUTO_TASK'
    $app.AutoTask()
    Start-Sleep -Milliseconds 500
    $taskCount=0;try{$taskCount=[int]$app.TasksList.Count}catch{}
    if($taskCount -le 0){throw 'SigmaNEST AutoTask completed but created no tasks.'}
    $phase='LABEL_AND_BATCH'
    $taskData=SN-Set-TaskNameAndBatch -app $app -requestParts $Request.parts
    $phase='SAVE'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'AutoTask'
    $ok=($taskData.warnings.Count -eq 0)
    return [pscustomobject]@{
      ok=$ok;phase='AUTOTASK_COMPLETE';wsPath=$wsPath;tasksCreated=$taskCount
      taskData=$taskData.tasks;warnings=$taskData.warnings;partUpdates=@($partUpdates)
      message=$(if($ok){'AutoTask created, labeled and batched '+$taskCount+' task(s).'}else{'AutoTask created '+$taskCount+' task(s) with warnings; see the Release Summary.'})
    }
  }catch{
    return [pscustomobject]@{
      ok=$false;phase=$phase;wsPath=$wsPath;tasksCreated=0;taskData=@();warnings=@($_.Exception.Message)
      message=('AutoTask failed at '+$phase+': '+$_.Exception.Message)
    }
  }finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}

function Invoke-SigmaNestBuild($Request){
  $phase='START';$app=$null
  try{
    $apartment=[Threading.Thread]::CurrentThread.GetApartmentState()
    if($apartment -ne [Threading.ApartmentState]::STA){throw ('Current PowerShell thread is '+$apartment+'; SigmaNEST COM requires STA.')}
    $phase='CREATE_COM';$app=New-Object -ComObject SigmaNEST.SNApp;if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $phase='RESOLVE_WS_PATH';$wsDir=[string]$Request.wsDirectory
    if([string]::IsNullOrWhiteSpace($wsDir)){
      $wsRoot=[string]$Request.wsRoot
      if([string]::IsNullOrWhiteSpace($wsRoot)){try{$parent=[IO.Directory]::GetParent([string]$Request.libraryRoot);if($parent){$wsRoot=$parent.FullName}}catch{}}
      if(-not [string]::IsNullOrWhiteSpace($wsRoot)){try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}}
    }
    if([string]::IsNullOrWhiteSpace($wsDir)){throw 'SigmaNEST WS output folder could not be determined. Configure the SigmaNEST WS folder (PathID 0) or verify the PRS library path.'}
    $wsDir=[IO.Path]::GetFullPath($wsDir.Trim());if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
    $job=[string]$Request.jobName;if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
    $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim();if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
    $wsPath=Join-Path -Path $wsDir -ChildPath ($safe+'.ws');if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
    $phase='CONFIGURE_LIBRARY';try{$app.PartsLibrary.Directory=[string]$Request.libraryRoot}catch{}
    $phase='IMPORT_PARTS';$created=@();$queued=@()
    $beforeParts=SN-Parts-Count $app
    foreach($x in @($Request.parts)){
      $sourcePath=[string]$x.sourcePath;$sourceType=[string]$x.sourceType
      if([string]::IsNullOrWhiteSpace($sourcePath)){$sourcePath=[string]$x.prsPath;if(-not $sourceType){$sourceType='PRS'}}
      if([string]::IsNullOrWhiteSpace($sourcePath)){continue}

      $load=SN-Queue-Geometry -app $app -sourcePath $sourcePath -sourceType $sourceType
      $quantity=SN-Scalar-Int -value $x.qty -default 1;if($quantity -lt 1){$quantity=1}
      $material=[string]$x.sigmaMaterial
      $thicknessText=$(if($x.thicknessMm -ne $null -and -not [double]::IsNaN([double]$x.thicknessMm)){[string]$x.thicknessMm}else{''})
      $queued += [pscustomobject]@{
        request=$x
        sourcePath=$sourcePath
        sourceType=$sourceType
        method=$load.method
      }
      $created += [pscustomobject]@{
        part=[string]$x.part
        qty=$quantity
        material=$material
        thickness=$thicknessText
        quantityProperty='TASK_PART_PENDING'
        materialProperty='TASK_PENDING'
        thicknessProperty='TASK_PENDING'
        sourcePath=$sourcePath
        sourceType=$sourceType
        sourceProperty=''
        batchMultiplier=$x.batchMultiplier
        taskBatches=@($x.taskBatches)
      }
    }

    if($queued.Count -eq 0){
      throw 'No geometry was queued for SigmaNEST import.'
    }

    $phase='COMMIT_IMPORTED_PARTS';$app.CreatePartsListForNewPartsInWS()
    $afterParts=SN-Parts-Count $app
    if($afterParts -lt ($beforeParts+$queued.Count)){
      throw ('SigmaNEST committed '+($afterParts-$beforeParts)+' part(s) but '+$queued.Count+' part(s) were requested for import.')
    }

    # Autosave the geometry workspace immediately after the imported parts are
    # committed. This guarantees a usable .ws exists even if later Task Setup
    # automation fails or disconnects a COM proxy.
    $phase='SAVE_GEOMETRY_CHECKPOINT'
    $saved=$false
    $saveErrors=@()
    for($attempt=1;$attempt -le 3 -and -not $saved;$attempt++){
      try{
        [void]$app.SaveWorkSpaceFile([string]$wsPath)
        Start-Sleep -Milliseconds 350
        $saved=Test-Path -LiteralPath $wsPath
      }catch{
        $saveErrors+=('Attempt '+$attempt+': '+(SN-ErrorText $_))
        if($attempt -lt 3){Start-Sleep -Milliseconds 500}
      }
    }
    if(-not $saved){
      throw ('SigmaNEST imported '+($afterParts-$beforeParts)+' part(s), but the automatic .ws save did not produce a file: '+$wsPath+'. '+($saveErrors -join ' | '))
    }

    $phase='APPLY_CL_PART_DATA';$partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts
    $phase='CREATE_TASKS';$app.CreateTasksListForNewPartsInWS()
    # Save again after the task list exists so the workspace remains resumable
    # if a later task attribute or quantity update fails.
    $phase='SAVE_TASK_CHECKPOINT'
    $savedTask=$false
    try{
      [void]$app.SaveWorkSpaceFile([string]$wsPath)
      Start-Sleep -Milliseconds 350
      $savedTask=Test-Path -LiteralPath $wsPath
    }catch{}
    if(-not $savedTask){
      throw ('SigmaNEST task list was created, but the automatic task checkpoint could not be confirmed on disk: '+$wsPath)
    }
    $phase='APPLY_TASK_QUANTITIES';SN-Set-TaskPartQuantity -app $app -requestParts $Request.parts
    $phase='SAVE_WS';$app.SaveWorkSpaceFile([string]$wsPath)
    try{$app.LoadWorkSpaceFile([string]$wsPath)}catch{};try{$app.RefreshTreeView()}catch{};try{$app.Redraw()}catch{}
    return [pscustomobject]@{ok=$true;creatorVersion='DIRECT-COM-1.8';phase='COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count;message=('SigmaNEST WS created: '+$wsPath)}
  }catch{
    $checkpointExists=$false
    try{$checkpointExists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false
      creatorVersion='DIRECT-COM-1.8'
      phase=$phase
      error=$_.Exception.Message
      category=$_.CategoryInfo.ToString()
      wsPath=$(if($checkpointExists){[string]$wsPath}else{''})
      parts=$created
      partCount=@($created).Count
      checkpointSaved=$checkpointExists
      message=$(if($checkpointExists){'SigmaNEST created a geometry/workspace checkpoint before the failure: '+[string]$wsPath}else{'SigmaNEST build failed before a workspace checkpoint was saved.'})
    }
  }
  finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
},'')
}
function SN-TaskPart-MatchesRequest($taskPart,$rp){
  if($null -eq $taskPart -or $null -eq $rp){return $false}
  $id=SN-Get-PartIdentity $taskPart
  $target=SN-Normalize-PartIdentity ([string]$rp.part)
  $source=SN-RequestSourceStem $rp
  if($target -and (
      $id.normalizedName -eq $target -or
      $id.normalizedDrawing -eq $target -or
      $id.normalizedPartNumber -eq $target
    )){return $true}
  if($source -and (
      $id.normalizedFile -eq $source -or
      $id.normalizedDrawing -eq $source -or
      $id.normalizedPartNumber -eq $source
    )){return $true}
  return $false
}
function SN-Find-RequestForTaskPart($taskPart,$requestParts){
  if($null -eq $taskPart){return $null}
  $id=SN-Get-PartIdentity $taskPart
  $rows=@()
  foreach($rp in @($requestParts)){
    $target=SN-Normalize-PartIdentity ([string]$rp.part)
    $source=SN-RequestSourceStem $rp
    $score=0
    if($target -and $id.normalizedName -eq $target){$score=100}
    elseif($target -and $id.normalizedPartNumber -eq $target){$score=98}
    elseif($target -and $id.normalizedDrawing -eq $target){$score=95}
    elseif($source -and $id.normalizedFile -eq $source){$score=90}
    elseif($source -and $id.normalizedDrawing -eq $source){$score=88}
    elseif($source -and $id.normalizedPartNumber -eq $source){$score=86}
    if($score -gt 0){$rows+=[pscustomobject]@{request=$rp;score=$score}}
  }
  if($rows.Count -eq 0){return $null}
  $top=@($rows|Sort-Object @{Expression={[int]$_.score};Descending=$true})
  if($top.Count -gt 1 -and [int]$top[0].score -eq [int]$top[1].score){return $null}
  $top[0].request
}
function SN-Find-WorkspacePartExact($app,$targetName,$sourcePath='',$usedIndices=@()){
  $targetNorm=SN-Normalize-PartIdentity $targetName
  $sourceStem=''
  try{$sourceStem=SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension([string]$sourcePath))}catch{}
  $sourceStemNoPrs=SN-Normalize-PartIdentity ($sourceStem -replace '(?i)PRS$','')
  $count=SN-Parts-Count $app
  $candidates=@()
  for($i=0;$i -lt $count;$i++){
    if(@($usedIndices) -contains $i){continue}
    $part=$null
    try{$part=$app.PartsList.Items($i)}catch{continue}
    if($null -eq $part){continue}
    $id=SN-Get-PartIdentity $part
    $score=0
    if($targetNorm -and $id.normalizedName -eq $targetNorm){$score=100}
    if($sourceStemNoPrs -and $id.normalizedFile -eq $sourceStemNoPrs){$score=[math]::Max($score,95)}
    if($targetNorm -and $id.normalizedDrawing -eq $targetNorm){$score=[math]::Max($score,90)}
    if($sourceStemNoPrs -and $id.normalizedDrawing -eq $sourceStemNoPrs){$score=[math]::Max($score,85)}
    if($score -gt 0){
      $candidates += [pscustomobject]@{part=$part;index=$i;score=$score;identity=$id}
    }
  }
  if($candidates.Count -eq 0){return $null}
  $top=@($candidates|Sort-Object @{Expression={[int]$_.score};Descending=$true},@{Expression={[int]$_.index};Descending=$false})
  if($top.Count -gt 1 -and [int]$top[0].score -eq [int]$top[1].score){
    # Prefer an exact normalized Name match over filename/drawing fallbacks.
    $exact=@($top|Where-Object {$_.identity.normalizedName -eq $targetNorm})
    if($exact.Count -eq 1){return $exact[0]}
    if($exact.Count -gt 1){throw ('Multiple SigmaNEST workspace parts match CL part "'+$targetName+'".')}
  }
  return $top[0]
}
function SN-Verify-PartIdentity($partObj,$targetName,$sourcePath=''){
  $id=SN-Get-PartIdentity $partObj
  $targetNorm=SN-Normalize-PartIdentity $targetName
  $sourceNorm=''
  try{$sourceNorm=SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension([string]$sourcePath))}catch{}
  $sourceNorm=SN-Normalize-PartIdentity ($sourceNorm -replace '(?i)PRS$','')
  if($targetNorm -and $id.normalizedName -eq $targetNorm){return $true}
  if($targetNorm -and $id.normalizedDrawing -eq $targetNorm){return $true}
  if($sourceNorm -and $id.normalizedFile -eq $sourceNorm){return $true}
  if($sourceNorm -and $id.normalizedDrawing -eq $sourceNorm){return $true}
  return $false
}
function SN-Apply-WorkspacePartData($app,$requestParts){
  $updated=@()
  $usedIndices=@()
  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $sourcePath=[string]$rp.sourcePath
    $found=SN-Find-WorkspacePartExact -app $app -targetName $targetName -sourcePath $sourcePath -usedIndices $usedIndices
    if($null -eq $found){
      throw ('Could not map imported SigmaNEST geometry back to CL part "'+$targetName+'". The imported PartsList order is not trusted.')
    }
    if(-not (SN-Verify-PartIdentity -partObj $found.part -targetName $targetName -sourcePath $sourcePath)){
      throw ('SigmaNEST workspace identity verification failed for CL part "'+$targetName+'". Refusing to overwrite another part.')
    }
    $usedIndices += [int]$found.index
    $partObj=$found.part
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $material=[string]$rp.sigmaMaterial
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    $row=[ordered]@{
      part=$targetName
      index=[int]$found.index
      qty=$qty
      material=$material
      thickness=$thickness
      quantityProperty=''
      materialProperty=''
      thicknessProperty=''
      warnings=@()
    }

    if(-not [string]::IsNullOrWhiteSpace($material)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @('Material','MaterialName','Mat','MatName','MaterialType') -value $material -expectedText $material
      if(-not $setInfo){
        throw ('CL material "'+$material+'" could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.materialProperty=[string]$setInfo.path
    }

    if(-not [double]::IsNaN($thickness)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -value $thickness -expectedNumber $thickness
      if(-not $setInfo){
        throw ('CL thickness '+$thickness+'mm could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.thicknessProperty=[string]$setInfo.path
    }

    $setInfo=SN-Set-PartField -partObj $partObj -names @('QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -value $qty -expectedInt $qty
    if(-not $setInfo){
      throw ('CL quantity '+$qty+' could not be written and verified on SigmaNEST part "'+$targetName+'".')
    }
    $row.quantityProperty=[string]$setInfo.path
    $updated += [pscustomobject]$row
  }
  @($updated)
}
function SN-Read-PartField($partObj,[string[]]$aliases,[string]$expectedText='',$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $partObj){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$partObj;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 4){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    foreach($prop in @(SN-ComPropertyNames $obj)){
      if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
      try{
        $v=$obj.$prop
        if($expectedText -ne ''){
          if(([string]$v).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=[string]$v}
          }
        }elseif(-not [double]::IsNaN($expectedNumber)){
          $n=SN-Scalar-Number -value $v -default ([double]::NaN)
          if(-not [double]::IsNaN([double]$n) -and [double]$n -eq $expectedNumber){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=$n}
          }
        }elseif($expectedInt -ne -2147483648){
          $n=SN-Scalar-Int -value $v -default -2147483648
          if($n -eq $expectedInt){
            return [pscustomobject]@{path=($node.path+'.'+$prop);value=$n}
          }
        }else{
          return [pscustomobject]@{path=($node.path+'.'+$prop);value=$v}
        }
      }catch{}
    }

    if($depth -ge 4){continue}
    foreach($prop in @(SN-ComPropertyNames $obj)){
      if($prop -in @('OwnerList','ParentObject','Parent','Application','Owner','Geometry','PartPolyLinesList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Verify-WorkspaceCLData($app,$requestParts){
  foreach($rp in @($requestParts)){
    $target=[string]$rp.part
    $source=[string]$rp.sourcePath
    $found=SN-Find-WorkspacePartExact -app $app -targetName $target -sourcePath $source -usedIndices @()
    if($null -eq $found){throw ('Post-save verification could not find SigmaNEST part "'+$target+'".')}
    $material=[string]$rp.sigmaMaterial
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','Mat','MatName','MaterialType') -expectedText $material
      if($null -eq $m){throw ('Post-save verification failed for "'+$target+'": material is not "'+$material+'".')}
    }
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    if(-not [double]::IsNaN($thickness)){
      $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -expectedNumber $thickness
      if($null -eq $t){throw ('Post-save verification failed for "'+$target+'": thickness is not '+$thickness+'mm.')}
    }
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $q=SN-Read-PartField -partObj $found.part -aliases @('QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $qty
    if($null -eq $q){throw ('Post-save verification failed for "'+$target+'": quantity is not '+$qty+'.')}
  }
  return $true
}
function SN-Parts-Count($app){try{return [int]$app.PartsList.Count}catch{return 0}}

function SN-Get-NewPart($app,$beforeCount,[string]$label){
  $beforeCount=SN-Scalar-Int -value $beforeCount -default 0;
  $afterCount=SN-Parts-Count $app
  if($afterCount -le $beforeCount){throw ('SigmaNEST loaded "'+$label+'" but did not add it to the workspace PartsList.')}
  try{$last=$null;foreach($item in $app.PartsList){$last=$item};if($null -ne $last){return $last}}catch{}
  foreach($member in @('Item','Items','get_Item')){
    try{$obj=$app.PartsList.GetType().InvokeMember($member,[Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::GetProperty,$null,$app.PartsList,@([int]($afterCount-1)));if($null -ne $obj){return $obj}}catch{}
  }
  throw ('SigmaNEST added "'+$label+'" but the new PartsList item could not be accessed. PartsList.Count='+$afterCount)
}

function SN-AllMemberDefinitions($obj){
  if($null -eq $obj){return @()}
  try{
    return @($obj | Get-Member -MemberType Methods,Properties,ParameterizedProperty,CodeProperty,NoteProperty,ScriptProperty,Fields -ErrorAction Stop |
      ForEach-Object {[pscustomobject]@{Name=[string]$_.Name;Type=[string]$_.MemberType;Definition=[string]$_.Definition}})
  }catch{return @()}
}
function SN-ComPropertyNamesAll($obj){
  return @((SN-AllMemberDefinitions $obj)|Where-Object {$_.Type -match 'Property'}|Select-Object -ExpandProperty Name -Unique)
}
function SN-InvokeSingleArgMember($obj,[string]$methodName,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $obj -or [string]::IsNullOrWhiteSpace($methodName)){return $null}
  try{
    $defs=@((SN-AllMemberDefinitions $obj)|Where-Object {$_.Name -eq $methodName})
    foreach($d in $defs){
      $def=[string]$d.Definition
      if($def -notmatch '\\(([^)]*)\\)'){continue}
      $inside=$Matches[1].Trim()
      if($inside -match ','){continue}
      if($inside -and $inside -notmatch '(?i)optional|paramarray'){ }
      try{
        $result=$obj.GetType().InvokeMember($methodName,[Reflection.BindingFlags]::InvokeMethod,$null,$obj,@($value))
        if($expectedText -ne ''){
          if($result -ne $null -and ([string]$result).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $methodName}
          try{
            $back=$obj.GetType().InvokeMember($methodName,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())
            if(([string]$back).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $methodName}
          }catch{}
        }elseif(-not [double]::IsNaN($expectedNumber)){
          try{if(([double]$result) -eq $expectedNumber){return $methodName}}catch{}
        }elseif($expectedInt -ne -2147483648){
          try{if((SN-Scalar-Int $result -default -2147483648) -eq $expectedInt){return $methodName}}catch{}
        }else{return $methodName}
      }catch{}
    }
  }catch{}
  return $null
}
function SN-Try-SetImportSetting($settings,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $settings){return $null}
  # 1) Standard writable/parameterized properties.
  foreach($prop in @(SN-ComPropertyNamesAll $settings)){
    if(-not (SN-FieldNameMatches -name $prop -aliases $aliases)){continue}
    try{
      $settings.$prop=$value
      if($expectedText -ne '' -and ([string]$settings.$prop).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      if(-not [double]::IsNaN($expectedNumber) -and (SN-Scalar-Number $settings.$prop -default ([double]::NaN)) -eq $expectedNumber){return $prop}
      if($expectedInt -ne -2147483648 -and (SN-Scalar-Int $settings.$prop -default -2147483648) -eq $expectedInt){return $prop}
    }catch{}
    try{
      $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::SetProperty,$null,$settings,@($value))|Out-Null
      if($expectedText -ne ''){
        $back=[string]$settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@())
        if($back.Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $prop}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        $back=SN-Scalar-Number $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@()) -default ([double]::NaN)
        if(-not [double]::IsNaN($back) -and $back -eq $expectedNumber){return $prop}
      }elseif($expectedInt -ne -2147483648){
        $back=SN-Scalar-Int $settings.GetType().InvokeMember($prop,[Reflection.BindingFlags]::GetProperty,$null,$settings,@()) -default -2147483648
        if($back -eq $expectedInt){return $prop}
      }
    }catch{}
  }

  # 2) Setter-style methods such as SetMaterial / SetThickness / SetQuantity.
  foreach($method in @(SN-AllMemberDefinitions $settings | Where-Object {$_.Type -eq 'Method'} | Select-Object -ExpandProperty Name -Unique)){
    $norm=SN-NormalizeFieldName $method
    foreach($alias in @($aliases)){
      $an=SN-NormalizeFieldName $alias
      if($norm -eq $an -or $norm -eq ('SET'+$an) -or $norm -eq ('SET'+$an+'VALUE') -or $norm -eq ('SET'+$an+'NAME') -or ($norm.StartsWith('SET') -and $norm.Contains($an))){
        $hit=SN-InvokeSingleArgMember -obj $settings -methodName $method -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
        if($hit){return ('METHOD.'+$hit)}
      }
    }
  }

  # 3) Generic field/value methods used by COM import-mapping objects.
  foreach($method in @('SetValue','SetField','SetParameter','SetProperty','SetImportValue','SetOption')){
    $defs=@((SN-AllMemberDefinitions $settings)|Where-Object {$_.Name -eq $method -and $_.Type -eq 'Method'})
    foreach($d in $defs){
      try{
        if(([string]$d.Definition) -match '\\([^)]*\\)' -and [string]$d.Definition -notmatch ','){
          $r=$settings.GetType().InvokeMember($method,[Reflection.BindingFlags]::InvokeMethod,$null,$settings,@($value))
          if($expectedText -ne '' -and $r -ne $null -and ([string]$r).Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return ('GENERIC.'+$method)}
          if($expectedInt -ne -2147483648 -and (SN-Scalar-Int $r -default -2147483648) -eq $expectedInt){return ('GENERIC.'+$method)}
          if(-not [double]::IsNaN($expectedNumber) -and (SN-Scalar-Number $r -default ([double]::NaN)) -eq $expectedNumber){return ('GENERIC.'+$method)}
        }
      }catch{}
    }
  }
  return $null
}
function SN-Get-ImportSettings($adapter){
  if($null -ne $adapter){
    foreach($prop in @('PartImportSettings','ImportSettings','Settings')){
      try{
        $v=$adapter.$prop
        if($null -ne $v){return [pscustomobject]@{settings=$v;property=$prop}}
      }catch{}
    }
  }
  foreach($progId in @('SigmaNEST.SNPartImportSettings','SigmaNEST.SNPartImportSetting')){
    try{
      $v=New-Object -ComObject $progId -ErrorAction Stop
      if($null -ne $v){return [pscustomobject]@{settings=$v;property=$progId}}
    }catch{}
  }
  return $null
}
function SN-Get-ImportDiagnostics($adapter,$settings){
  $parts=@()
  if($null -ne $adapter){
    $parts+='Adapter methods/properties:'
    foreach($d in @(SN-AllMemberDefinitions $adapter)){$parts+=('  '+$d.Name+' :: '+$d.Definition)}
  }
  if($null -ne $settings){
    $parts+='PartImportSettings methods/properties:'
    foreach($d in @(SN-AllMemberDefinitions $settings)){$parts+=('  '+$d.Name+' :: '+$d.Definition)}
    foreach($prop in @(SN-ComPropertyNamesAll $settings)){
      try{
        $child=$settings.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array] -and $child.Count -gt 1){continue}
        $parts+=('  Nested '+$prop+':')
        foreach($d in @(SN-AllMemberDefinitions $child)){$parts+=('    '+$d.Name+' :: '+$d.Definition)}
      }catch{}
    }
  }
  return @($parts)
}
function SN-SetImportSettingDeep($root,[string[]]$aliases,$value,[string]$expectedText='',[double]$expectedNumber=([double]::NaN),$expectedInt=-2147483648){
  if($null -eq $root){return $null}
  $visited=@{}
  $queue=New-Object System.Collections.Queue
  $queue.Enqueue([pscustomobject]@{object=$root;path='ROOT';depth=0})
  while($queue.Count -gt 0){
    $node=$queue.Dequeue()
    $obj=$node.object;$depth=[int]$node.depth
    if($null -eq $obj -or $depth -gt 5){continue}
    $identity=''
    try{$identity=[string][Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($obj)}catch{}
    if($identity -and $visited.ContainsKey($identity)){continue}
    if($identity){$visited[$identity]=$true}

    $hit=SN-Try-SetImportSetting -settings $obj -aliases $aliases -value $value -expectedText $expectedText -expectedNumber $expectedNumber -expectedInt $expectedInt
    if($hit){return [pscustomobject]@{path=($node.path+'.'+$hit)}}

    if($depth -ge 5){continue}
    foreach($prop in @(SN-ComPropertyNamesAll $obj)){
      if($prop -in @('Owner','Parent','ParentObject','Application','PartImportSettings','SolidsList','Count','Item','Items')){continue}
      try{
        $child=$obj.$prop
        if($null -eq $child -or $child -is [string] -or $child -is [ValueType]){continue}
        if($child -is [System.Array]){
          if($child.Count -eq 1){$child=$child[0]}else{continue}
        }
        $queue.Enqueue([pscustomobject]@{object=$child;path=($node.path+'.'+$prop);depth=$depth+1})
      }catch{}
    }
  }
  return $null
}
function SN-Configure-PRS-ImportSettings($settings,$clData){
  $material=[string]$clData.sigmaMaterial
  $thickness=SN-Scalar-Number -value $clData.thicknessMm -default ([double]::NaN)
  $qty=SN-Scalar-Int -value $clData.qty -default 1
  if($qty -lt 1){$qty=1}
  $result=[ordered]@{material='';thickness='';quantity='';warnings=@()}
  if(-not [string]::IsNullOrWhiteSpace($material)){
    $r=SN-SetImportSettingDeep -root $settings -aliases @('Material','MaterialName','PartMaterial','MaterialType','MaterialString') -value $material -expectedText $material
    if($r){$result.material=$r.path}
  }
  if(-not [double]::IsNaN($thickness)){
    $r=SN-SetImportSettingDeep -root $settings -aliases @('Thickness','SheetThickness','MaterialThickness','MaterialThk','Thk','Thick') -value $thickness -expectedNumber $thickness
    if($r){$result.thickness=$r.path}
  }
  $r=SN-SetImportSettingDeep -root $settings -aliases @('QtyOrdered','Quantity','PartQuantity','NumberToNest','NumberToLoad','QuantityToNest','QtyToNest','BatchQuantity') -value $qty -expectedInt $qty
  if($r){$result.quantity=$r.path}
  if([string]::IsNullOrWhiteSpace([string]$result.material)){ $result.warnings+='Import settings did not expose a verified material field.' }
  if(-not [double]::IsNaN($thickness) -and [string]::IsNullOrWhiteSpace([string]$result.thickness)){ $result.warnings+='Import settings did not expose a verified thickness field.' }
  if([string]::IsNullOrWhiteSpace([string]$result.quantity)){ $result.warnings+='Import settings did not expose a verified quantity field.' }
  return [pscustomobject]$result
}
function SN-Try-ImportPRS-WithSettings($app,[string]$sourcePath,$clData){
  $before=SN-Parts-Count $app
  $adapter=$null
  $info=$null
  $diagnostics=@()
  try{$adapter=New-Object -ComObject SigmaNEST.SNPartExportImport -ErrorAction Stop}catch{
    $diagnostics+='SNPartExportImport unavailable: '+(SN-ErrorText $_)
  }
  if($null -ne $adapter){
    $info=SN-Get-ImportSettings -adapter $adapter
    if($info){
      $cfg=SN-Configure-PRS-ImportSettings -settings $info.settings -clData $clData
      if($cfg.warnings.Count -eq 0){
        foreach($methodName in @('ImportPartWithFeedback','ImportPart2','ImportPart','ImportAsParts')){
          $members=@(SN-ComMethodNames $adapter)
          if($members -notcontains $methodName){continue}
          foreach($args in @(
            @([string]$sourcePath,$info.settings),
            @([string]$sourcePath,$info.settings,$false),
            @([string]$sourcePath)
          )){
            try{
              $result=SN-Invoke-ComMethod -obj $adapter -name $methodName -args $args
              Start-Sleep -Milliseconds 150
              $after=SN-Parts-Count $app
              if($after -gt $before){
                return [pscustomobject]@{ok=$true;method=('SNPartExportImport.'+$methodName);settings=$cfg}
              }
              if($result -is [System.Object[]]){
                $result=@($result|Where-Object {$null -ne $_})[0]
              }
              if($result -and $result.PSObject.Properties.Name -contains 'Part'){
                try{
                  $part=$result.Part
                  if($null -ne $part){return [pscustomobject]@{ok=$true;method=('SNPartExportImport.'+$methodName);settings=$cfg;pendingPart=$part}}
                }catch{}
              }
            }catch{
              $diagnostics+=($methodName+' '+$args.Count+' args: '+(SN-ErrorText $_))
            }
          }
        }
      }else{
        $diagnostics+=($cfg.warnings -join ' ')
      }
    }else{
      $diagnostics+='SNPartExportImport has no accessible PartImportSettings/ImportSettings/Settings object.'
    }
  }
  if($null -ne $adapter){
    try{$diagnostics+='Adapter methods: '+((SN-ComMethodNames $adapter) -join ', ')}catch{}
    try{$diagnostics+='Adapter properties: '+((SN-ComPropertyNames $adapter) -join ', ')}catch{}
  }
  try{if($adapter){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($adapter)}}catch{}
  if($null -ne $adapter -and $null -ne $info -and $null -ne $info.settings){
    $diagnostics += @(SN-Get-ImportDiagnostics -adapter $adapter -settings $info.settings)
  }
  return [pscustomobject]@{ok=$false;diagnostics=$diagnostics}
}
function SN-Queue-Geometry($app,[string]$sourcePath,[string]$sourceType,$clData=$null){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $errors=@()

  if($sourceType -eq 'PRS' -and $null -ne $clData){
    $prsTry=SN-Try-ImportPRS-WithSettings -app $app -sourcePath $sourcePath -clData $clData
    if($prsTry.ok){
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method=$prsTry.method;preImportSettings=$prsTry.settings}
    }
    if($prsTry.diagnostics.Count){$errors+=($prsTry.diagnostics -join ' | ')}
  }

  try{
    # LoadPart is retained as a geometry-loader fallback for DXF and for
    # installations where PRS import is known to work. PRS is NOT allowed
    # to fall through silently when CL import-time settings were unavailable.
    if($sourceType -ne 'PRS'){
      [void]$app.LoadPart([string]$sourcePath)
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method='LoadPart'}
    }
  }catch{
    $errors+=('LoadPart: '+(SN-ErrorText $_))
  }

  if($sourceType -eq 'DXF'){
    try{
      $settings=$null;try{$settings=New-Object -ComObject SigmaNEST.SNPartImportSettings}catch{}
      foreach($importId in @(0,1,2,3)){
        try{
          $result=SN-Invoke-ComMethod $app.PartsList 'Import' @([string]$sourcePath,[int]$importId,$settings)
          return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method=('PartsList.Import '+$importId)}
        }catch{
          $errors+=('PartsList.Import id '+$importId+': '+(SN-ErrorText $_))
        }
      }
    }catch{
      $errors+=('DXF import fallback: '+(SN-ErrorText $_))
    }
  }

  if($sourceType -eq 'PRS'){
    throw ('SigmaNEST could not perform a verified CL-data PRS import for "'+$label+'". The PRS was NOT imported because its original material/thickness/quantity cannot be safely accepted as production data. '+($errors -join ' | '))
  }
  throw ('SigmaNEST could not load "'+$label+'". '+($errors -join ' | '))
}


function SN-Set-TaskMaterialAndThickness($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){
    throw 'SigmaNEST created no TasksList entries after CreateTasksListForNewPartsInWS; cannot apply CL material/thickness safely.'
  }

  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $material=[string]$rp.sigmaMaterial
    $thickness=$null
    try{
      if($rp.thicknessMm -ne $null -and -not [double]::IsNaN([double]$rp.thicknessMm)){
        $thickness=[double]$rp.thicknessMm
      }
    }catch{}

    $matched=$false
    $matchedTaskIndex=-1
    $matchedPartIndex=-1

    for($ti=0;$ti -lt $taskCount -and -not $matched;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}

      $taskPartCount=0
      try{$taskPartCount=[int]$task.PartsList.Count}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}
        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not $taskPartName.Equals($targetName,[StringComparison]::OrdinalIgnoreCase)){continue}

        $matched=$true
        $matchedTaskIndex=$ti
        $matchedPartIndex=$pi
        break
      }
    }

    if(-not $matched){
      throw ('Could not find CL part "'+$targetName+'" inside any SigmaNEST task after task creation.')
    }

    # Reacquire the COM task/part immediately before each mutation. SigmaNEST
    # may rebuild the task tree when Task Setup data changes, invalidating
    # previously returned COM interfaces.
    if(-not [string]::IsNullOrWhiteSpace($material)){
      $task=$null;$taskPart=$null
      try{$task=$app.TasksList.Items($matchedTaskIndex)}catch{}
      if($null -eq $task){throw ('Could not reacquire SigmaNEST task index '+$matchedTaskIndex+' for part "'+$targetName+'".')}
      try{$taskPart=$task.PartsList.Items($matchedPartIndex)}catch{}
      if($null -eq $taskPart){throw ('Could not reacquire SigmaNEST task-part index '+$matchedPartIndex+' for part "'+$targetName+'".')}

      $setMaterial=$false
      foreach($propertyName in @('Material','MaterialName','Mat')){
        try{
          $taskPart.$propertyName=$material
          $readBack=[string]$taskPart.$propertyName
          if($readBack.Trim().Equals($material.Trim(),[StringComparison]::OrdinalIgnoreCase)){
            $setMaterial=$true
            break
          }
        }catch{
          if(SN-IsDisconnected $_){
            throw ('SigmaNEST COM task-part disconnected while setting material property "'+$propertyName+'" for CL part "'+$targetName+'". HRESULT 0x80010108 (RPC_E_DISCONNECTED).')
          }
        }
      }

      if(-not $setMaterial){
        throw ('SigmaNEST task-part for "'+$targetName+'" does not expose a writable material property. Tried: Material, MaterialName, Mat.')
      }
    }

    if($null -ne $thickness){
      $task=$null;$taskPart=$null
      try{$task=$app.TasksList.Items($matchedTaskIndex)}catch{}
      if($null -eq $task){throw ('Could not reacquire SigmaNEST task index '+$matchedTaskIndex+' for part "'+$targetName+'" before thickness update.')}
      try{$taskPart=$task.PartsList.Items($matchedPartIndex)}catch{}
      if($null -eq $taskPart){throw ('Could not reacquire SigmaNEST task-part index '+$matchedPartIndex+' for part "'+$targetName+'" before thickness update.')}

      $setThickness=$false
      foreach($propertyName in @('Thickness','SheetThickness','Thk','MaterialThickness')){
        try{
          $taskPart.$propertyName=$thickness
          $readBack=[double]$taskPart.$propertyName
          if($readBack -eq $thickness){
            $setThickness=$true
            break
          }
        }catch{
          if(SN-IsDisconnected $_){
            throw ('SigmaNEST COM task-part disconnected while setting thickness property "'+$propertyName+'" for CL part "'+$targetName+'". HRESULT 0x80010108 (RPC_E_DISCONNECTED).')
          }
        }
      }

      if(-not $setThickness){
        throw ('SigmaNEST task-part for "'+$targetName+'" does not expose a writable thickness property. Tried: Thickness, SheetThickness, Thk, MaterialThickness.')
      }
    }
  }
}

function SN-Set-TaskPartQuantity($app,$requestParts){
  $taskCount=0
  try{$taskCount=[int]$app.TasksList.Count}catch{}
  if($taskCount -le 0){
    throw 'SigmaNEST created no TasksList entries after CreateTasksListForNewPartsInWS; cannot apply CL quantities safely.'
  }

  foreach($rp in @($requestParts)){
    $targetName=[string]$rp.part
    $quantity=SN-Scalar-Int -value $rp.qty -default 1
    if($quantity -lt 1){$quantity=1}
    $matched=$false

    for($ti=0;$ti -lt $taskCount -and -not $matched;$ti++){
      $task=$null
      try{$task=$app.TasksList.Items($ti)}catch{continue}
      if($null -eq $task){continue}

      $taskPartCount=0
      try{$taskPartCount=[int]$task.PartsList.Count}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount -and -not $matched;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}

        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not $taskPartName.Equals($targetName,[StringComparison]::OrdinalIgnoreCase)){continue}

        $set=$null
        $readBack=$null
        foreach($propertyName in @('Quantity','Qty','QtyRequired','QtyReq','BatchQty')){
          try{
            $taskPart.$propertyName=$quantity
            $readBack=[double]$taskPart.$propertyName
            if($readBack -eq $quantity){
              $set=$propertyName
              break
            }
          }catch{}
        }

        if(-not $set){
          throw ('SigmaNEST task part "'+$taskPartName+'" does not expose a writable quantity property. Task index='+$ti+'.')
        }

        $matched=$true
      }
    }

    if(-not $matched){
      throw ('Could not find CL part "'+$targetName+'" inside any SigmaNEST task after task creation.')
    }
  }
}
function SN-Resolve-WS-Path($Request){
  $wsDir=[string]$Request.wsDirectory
  if([string]::IsNullOrWhiteSpace($wsDir)){
    $wsRoot=[string]$Request.wsRoot
    if([string]::IsNullOrWhiteSpace($wsRoot)){
      try{
        $parent=[IO.Directory]::GetParent([string]$Request.libraryRoot)
        if($parent){$wsRoot=$parent.FullName}
      }catch{}
    }
    if(-not [string]::IsNullOrWhiteSpace($wsRoot)){
      try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}
    }
  }
  if([string]::IsNullOrWhiteSpace($wsDir)){throw 'SigmaNEST WS output folder could not be determined.'}
  $wsDir=[IO.Path]::GetFullPath($wsDir.Trim())
  if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
  return $wsDir
}

function SN-Save-WorkspaceVerified($app,[string]$wsPath,[string]$label){
  $errors=@()
  for($attempt=1;$attempt -le 3;$attempt++){
    try{
      [void]$app.SaveWorkSpaceFile([string]$wsPath)
      Start-Sleep -Milliseconds 350
      if(Test-Path -LiteralPath $wsPath){
        return [pscustomobject]@{ok=$true;attempt=$attempt;path=$wsPath;label=$label}
      }
    }catch{
      $errors+=('Attempt '+$attempt+': '+(SN-ErrorText $_))
    }
    if($attempt -lt 3){Start-Sleep -Milliseconds 500}
  }
  throw ('Automatic '+$label+' workspace save failed: '+$wsPath+'. '+($errors -join ' | '))
}

function Invoke-SigmaNestImportGeometry($Request){
  $phase='START';$app=$null;$created=@();$partUpdates=@()
  try{
    if([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA){throw 'SigmaNEST COM requires STA.'}
    $app=New-Object -ComObject SigmaNEST.SNApp
    if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $wsDir=SN-Resolve-WS-Path -Request $Request
    $job=[string]$Request.jobName
    if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
    $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim()
    if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
    $wsPath=Join-Path $wsDir ($safe+'.ws')
    if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
    try{$app.PartsLibrary.Directory=[string]$Request.libraryRoot}catch{}
    $phase='IMPORT_PARTS'
    $before=SN-Parts-Count $app
    $queued=@()
    $queuedIndex=0
    foreach($x in @($Request.parts)){
      $source=[string]$x.sourcePath
      if([string]::IsNullOrWhiteSpace($source)){$source=[string]$x.prsPath}
      if([string]::IsNullOrWhiteSpace($source)){continue}
      $clData=[pscustomobject]@{
        sigmaMaterial=[string]$x.sigmaMaterial
        thicknessMm=$x.thicknessMm
        qty=$x.qty
      }
      $load=SN-Queue-Geometry -app $app -sourcePath $source -sourceType ([string]$x.sourceType) -clData $clData
      $qty=SN-Scalar-Int -value $x.qty -default 1
      if($qty -lt 1){$qty=1}
      $workspaceIndex=$before+$queuedIndex
      $created+=[pscustomobject]@{
        part=[string]$x.part;qty=$qty;material=[string]$x.sigmaMaterial
        thickness=SN-Scalar-Number -value $x.thicknessMm -default ([double]::NaN)
        sourcePath=$source;sourceType=[string]$x.sourceType
        matchType=[string]$x.matchType;batchMultiplier=SN-Scalar-Int -value $x.batchMultiplier -default 1
        workspaceIndex=$workspaceIndex
      }
      $queued += [pscustomobject]@{
        part=[string]$x.part;qty=$qty;sigmaMaterial=[string]$x.sigmaMaterial
        thicknessMm=$x.thicknessMm;sourcePath=$source;sourceType=[string]$x.sourceType
      }
      $queuedIndex++
    }
    if($queued.Count -eq 0){throw 'No geometry was found to import.'}
    $phase='COMMIT_IMPORTED_PARTS'
    $app.CreatePartsListForNewPartsInWS()
    $after=SN-Parts-Count $app
    if($after -lt ($before+$queued.Count)){throw ('SigmaNEST committed '+($after-$before)+' part(s) but '+$queued.Count+' were requested.')}
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts
    $phase='SAVE_GEOMETRY'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'geometry'
    $phase='VERIFY_SAVED_CL_DATA'
    $app.LoadWorkSpaceFile([string]$wsPath)
    $verify=SN-Verify-WorkspaceCLData -app $app -requestParts $queued
    return [pscustomobject]@{
      ok=$true;phase='IMPORT_COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count
      partUpdates=@($partUpdates);tasksCreated=0;message=('Geometry imported and saved to '+$wsPath)
      sourceCount=$created.Count
    }
  }catch{
    $exists=$false;try{$exists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false;phase=$phase;wsPath=$(if($exists){[string]$wsPath}else{''})
      parts=$created;partCount=@($created).Count;partUpdates=@($partUpdates)
      checkpointSaved=$exists;error=$_.Exception.Message
      message=$(if($exists){'Geometry workspace saved before failure: '+$wsPath}else{'Geometry import failed before a verified workspace save.'})
    }
  }finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}

function SN-Set-TaskNameAndBatch($app,$requestParts){
  $taskCount=0;try{$taskCount=[int]$app.TasksList.Count}catch{}
  $map=@{};foreach($rp in @($requestParts)){$map[[string]$rp.part]=$rp}
  $results=@();$warnings=@()
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$app.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $names=@();$materials=@();$thks=@();$multis=@()
    $pc=0;try{$pc=[int]$task.PartsList.Count}catch{}
    for($pi=0;$pi -lt $pc;$pi++){
      try{$tp=$task.PartsList.Items($pi)}catch{continue}
      $pn='';try{$pn=[string]$tp.Name}catch{}
      if($map.ContainsKey($pn)){
        $rp=$map[$pn];$names+=$pn
        if(-not [string]::IsNullOrWhiteSpace([string]$rp.sigmaMaterial)){$materials+=[string]$rp.sigmaMaterial}
        $th=[double](SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN))
        if(-not [double]::IsNaN($th)){$thks+=$th}
        foreach($tb in @($rp.taskBatches)){
          $multis+=SN-Scalar-Int -value $tb.batchMultiplier -default 1
        }
        if($multis.Count -eq 0){$multis+=SN-Scalar-Int -value $rp.batchMultiplier -default 1}
      }
    }
    if($names.Count -eq 0){continue}
    $mat=if($materials.Count){$materials[0]}else{'UNKNOWN MATERIAL'}
    $thk=if($thks.Count){$thks[0]}else{0}
    $uniqueMulti=@($multis|Sort-Object -Unique)
    $label=('{0:00} | {1} | {2:0.###}mm' -f ($ti+1),$mat,$thk)
    $labelApplied=$false;$labelProp=''
    foreach($prop in @('Name','TaskName','Description')){
      try{
        $task.$prop=$label
        $back=[string]$task.$prop
        if($back.Trim().Equals($label.Trim(),[StringComparison]::OrdinalIgnoreCase)){$labelApplied=$true;$labelProp=$prop;break}
      }catch{}
    }
    if(-not $labelApplied){$warnings+=('Task '+($ti+1)+' could not be renamed; material/thickness grouping still exists.')}
    $batchApplied=$false;$batchProp='';$batch=1
    if($uniqueMulti.Count -eq 1){
      $batch=SN-Scalar-Int -value $uniqueMulti[0] -default 1
      if($batch -lt 1){$batch=1}
      foreach($prop in @('BatchMultiplier','BatchQty','BatchQuantity','Batch')){
        try{
          $task.$prop=$batch
          $back=SN-Scalar-Int -value $task.$prop -default -1
          if($back -eq $batch){$batchApplied=$true;$batchProp=$prop;break}
        }catch{}
      }
      if(-not $batchApplied){
        foreach($pi in 0..([math]::Max(0,$pc-1))){
          try{$tp=$task.PartsList.Items($pi)}catch{continue}
          foreach($prop in @('BatchQty','BatchMultiplier','BatchQuantity','Batch')){
            try{
              $tp.$prop=$batch
              $back=SN-Scalar-Int -value $tp.$prop -default -1
              if($back -eq $batch){$batchApplied=$true;$batchProp='PART.'+$prop;break}
            }catch{}
          }
          if($batchApplied){break}
        }
      }
    }else{
      $warnings+=('Task '+($ti+1)+' has mixed CL multipliers: '+($uniqueMulti -join ', ')+'. No multiplier was guessed.')
    }
    if(-not $batchApplied -and $uniqueMulti.Count -eq 1){
      $warnings+=('Task '+($ti+1)+' could not accept Batch multiplier x'+$batch+'.')
    }
    $results+=[pscustomobject]@{
      taskIndex=$ti+1;taskName=$label;material=$mat;thickness=$thk;batchMultiplier=$batch
      batchApplied=$batchApplied;batchProperty=$batchProp;labelApplied=$labelApplied;labelProperty=$labelProp
      partCount=$names.Count;multipliers=($uniqueMulti -join ',')
    }
  }
  return [pscustomobject]@{tasks=@($results);warnings=@($warnings)}
}

function Invoke-SigmaNestAutoTask($Request){
  $phase='START';$app=$null
  try{
    if([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA){throw 'SigmaNEST COM requires STA.'}
    $wsPath=[IO.Path]::GetFullPath([string]$Request.wsPath)
    if(-not(Test-Path -LiteralPath $wsPath)){throw ('SigmaNEST WS not found: '+$wsPath)}
    $app=New-Object -ComObject SigmaNEST.SNApp
    if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $phase='LOAD_WORKSPACE'
    $app.LoadWorkSpaceFile([string]$wsPath)
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $queued
    $phase='AUTO_TASK'
    $app.AutoTask()
    Start-Sleep -Milliseconds 500
    $taskCount=0;try{$taskCount=[int]$app.TasksList.Count}catch{}
    if($taskCount -le 0){throw 'SigmaNEST AutoTask completed but created no tasks.'}
    $phase='LABEL_AND_BATCH'
    $taskData=SN-Set-TaskNameAndBatch -app $app -requestParts $Request.parts
    $phase='SAVE'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'AutoTask'
    $ok=($taskData.warnings.Count -eq 0)
    return [pscustomobject]@{
      ok=$ok;phase='AUTOTASK_COMPLETE';wsPath=$wsPath;tasksCreated=$taskCount
      taskData=$taskData.tasks;warnings=$taskData.warnings;partUpdates=@($partUpdates)
      message=$(if($ok){'AutoTask created, labeled and batched '+$taskCount+' task(s).'}else{'AutoTask created '+$taskCount+' task(s) with warnings; see the Release Summary.'})
    }
  }catch{
    return [pscustomobject]@{
      ok=$false;phase=$phase;wsPath=$wsPath;tasksCreated=0;taskData=@();warnings=@($_.Exception.Message)
      message=('AutoTask failed at '+$phase+': '+$_.Exception.Message)
    }
  }finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}

function Invoke-SigmaNestBuild($Request){
  $phase='START';$app=$null
  try{
    $apartment=[Threading.Thread]::CurrentThread.GetApartmentState()
    if($apartment -ne [Threading.ApartmentState]::STA){throw ('Current PowerShell thread is '+$apartment+'; SigmaNEST COM requires STA.')}
    $phase='CREATE_COM';$app=New-Object -ComObject SigmaNEST.SNApp;if($null -eq $app){throw 'SigmaNEST.SNApp returned null.'}
    $phase='RESOLVE_WS_PATH';$wsDir=[string]$Request.wsDirectory
    if([string]::IsNullOrWhiteSpace($wsDir)){
      $wsRoot=[string]$Request.wsRoot
      if([string]::IsNullOrWhiteSpace($wsRoot)){try{$parent=[IO.Directory]::GetParent([string]$Request.libraryRoot);if($parent){$wsRoot=$parent.FullName}}catch{}}
      if(-not [string]::IsNullOrWhiteSpace($wsRoot)){try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}}
    }
    if([string]::IsNullOrWhiteSpace($wsDir)){throw 'SigmaNEST WS output folder could not be determined. Configure the SigmaNEST WS folder (PathID 0) or verify the PRS library path.'}
    $wsDir=[IO.Path]::GetFullPath($wsDir.Trim());if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
    $job=[string]$Request.jobName;if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
    $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim();if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
    $wsPath=Join-Path -Path $wsDir -ChildPath ($safe+'.ws');if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
    $phase='CONFIGURE_LIBRARY';try{$app.PartsLibrary.Directory=[string]$Request.libraryRoot}catch{}
    $phase='IMPORT_PARTS';$created=@();$queued=@()
    $beforeParts=SN-Parts-Count $app
    foreach($x in @($Request.parts)){
      $sourcePath=[string]$x.sourcePath;$sourceType=[string]$x.sourceType
      if([string]::IsNullOrWhiteSpace($sourcePath)){$sourcePath=[string]$x.prsPath;if(-not $sourceType){$sourceType='PRS'}}
      if([string]::IsNullOrWhiteSpace($sourcePath)){continue}

      $load=SN-Queue-Geometry -app $app -sourcePath $sourcePath -sourceType $sourceType
      $quantity=SN-Scalar-Int -value $x.qty -default 1;if($quantity -lt 1){$quantity=1}
      $material=[string]$x.sigmaMaterial
      $thicknessText=$(if($x.thicknessMm -ne $null -and -not [double]::IsNaN([double]$x.thicknessMm)){[string]$x.thicknessMm}else{''})
      $queued += [pscustomobject]@{
        request=$x
        sourcePath=$sourcePath
        sourceType=$sourceType
        method=$load.method
      }
      $created += [pscustomobject]@{
        part=[string]$x.part
        qty=$quantity
        material=$material
        thickness=$thicknessText
        quantityProperty='TASK_PART_PENDING'
        materialProperty='TASK_PENDING'
        thicknessProperty='TASK_PENDING'
        sourcePath=$sourcePath
        sourceType=$sourceType
        sourceProperty=''
        batchMultiplier=$x.batchMultiplier
        taskBatches=@($x.taskBatches)
      }
    }

    if($queued.Count -eq 0){
      throw 'No geometry was queued for SigmaNEST import.'
    }

    $phase='COMMIT_IMPORTED_PARTS';$app.CreatePartsListForNewPartsInWS()
    $afterParts=SN-Parts-Count $app
    if($afterParts -lt ($beforeParts+$queued.Count)){
      throw ('SigmaNEST committed '+($afterParts-$beforeParts)+' part(s) but '+$queued.Count+' part(s) were requested for import.')
    }

    # Autosave the geometry workspace immediately after the imported parts are
    # committed. This guarantees a usable .ws exists even if later Task Setup
    # automation fails or disconnects a COM proxy.
    $phase='SAVE_GEOMETRY_CHECKPOINT'
    $saved=$false
    $saveErrors=@()
    for($attempt=1;$attempt -le 3 -and -not $saved;$attempt++){
      try{
        [void]$app.SaveWorkSpaceFile([string]$wsPath)
        Start-Sleep -Milliseconds 350
        $saved=Test-Path -LiteralPath $wsPath
      }catch{
        $saveErrors+=('Attempt '+$attempt+': '+(SN-ErrorText $_))
        if($attempt -lt 3){Start-Sleep -Milliseconds 500}
      }
    }
    if(-not $saved){
      throw ('SigmaNEST imported '+($afterParts-$beforeParts)+' part(s), but the automatic .ws save did not produce a file: '+$wsPath+'. '+($saveErrors -join ' | '))
    }

    $phase='APPLY_CL_PART_DATA';$partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts
    $phase='CREATE_TASKS';$app.CreateTasksListForNewPartsInWS()
    # Save again after the task list exists so the workspace remains resumable
    # if a later task attribute or quantity update fails.
    $phase='SAVE_TASK_CHECKPOINT'
    $savedTask=$false
    try{
      [void]$app.SaveWorkSpaceFile([string]$wsPath)
      Start-Sleep -Milliseconds 350
      $savedTask=Test-Path -LiteralPath $wsPath
    }catch{}
    if(-not $savedTask){
      throw ('SigmaNEST task list was created, but the automatic task checkpoint could not be confirmed on disk: '+$wsPath)
    }
    $phase='APPLY_TASK_QUANTITIES';SN-Set-TaskPartQuantity -app $app -requestParts $Request.parts
    $phase='SAVE_WS';$app.SaveWorkSpaceFile([string]$wsPath)
    try{$app.LoadWorkSpaceFile([string]$wsPath)}catch{};try{$app.RefreshTreeView()}catch{};try{$app.Redraw()}catch{}
    return [pscustomobject]@{ok=$true;creatorVersion='DIRECT-COM-1.8';phase='COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count;message=('SigmaNEST WS created: '+$wsPath)}
  }catch{
    $checkpointExists=$false
    try{$checkpointExists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false
      creatorVersion='DIRECT-COM-1.8'
      phase=$phase
      error=$_.Exception.Message
      category=$_.CategoryInfo.ToString()
      wsPath=$(if($checkpointExists){[string]$wsPath}else{''})
      parts=$created
      partCount=@($created).Count
      checkpointSaved=$checkpointExists
      message=$(if($checkpointExists){'SigmaNEST created a geometry/workspace checkpoint before the failure: '+[string]$wsPath}else{'SigmaNEST build failed before a workspace checkpoint was saved.'})
    }
  }
  finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}