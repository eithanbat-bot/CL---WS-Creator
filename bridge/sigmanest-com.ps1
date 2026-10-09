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
        $back=SN-Scalar-Number -value $obj.$name -default ([double]::NaN)
        if(-not [double]::IsNaN([double]$back) -and $back -eq $expectedNumber){return $name}
      }elseif($expectedInt -ne -2147483648){
        $back=SN-Scalar-Int -value $obj.$name -default -2147483648
        if($back -eq $expectedInt){return $name}
      }else{return $name}
    }catch{}
    try{
      $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::SetProperty,$null,$obj,@($value))|Out-Null
      if($expectedText -ne ''){
        $back=[string]$obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$obj,@())
        if($back.Trim().Equals($expectedText.Trim(),[StringComparison]::OrdinalIgnoreCase)){return $name}
      }elseif(-not [double]::IsNaN($expectedNumber)){
        try{
          $back=SN-Scalar-Number -value $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$obj,@()) -default ([double]::NaN)
          if(-not [double]::IsNaN([double]$back) -and $back -eq $expectedNumber){return $name}
        }catch{}
      }elseif($expectedInt -ne -2147483648){
        try{
          $back=SN-Scalar-Int -value $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::GetProperty,$null,$obj,@()) -default -2147483648
          if($back -eq $expectedInt){return $name}
        }catch{}
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
function SN-Set-QtyToNest($partObj,[int]$qty){
  if($null -eq $partObj){return $null}
  try{
    $partObj.QtyToNest=$qty
    $back=SN-Scalar-Int -value $partObj.QtyToNest -default -2147483648
    if($back -eq $qty){return 'QtyToNest'}
  }catch{}
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
  if($null -eq $partObj){return [pscustomobject]@{name='';drawing='';file='';normalizedName='';normalizedDrawing='';normalizedFile=''}}
  $name='';$drawing='';$file=''
  try{$name=[string]$partObj.Name}catch{}
  try{$drawing=[string]$partObj.DrawingNumber}catch{}
  try{$file=[string]$partObj.PartFilename}catch{}
  return [pscustomobject]@{
    name=$name
    drawing=$drawing
    file=$file
    normalizedName=(SN-Normalize-PartIdentity $name)
    normalizedDrawing=(SN-Normalize-PartIdentity $drawing)
    normalizedFile=(SN-Normalize-PartIdentity ([IO.Path]::GetFileNameWithoutExtension($file)))
  }
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
function SN-Make-CLLinkKey($jobName,$rp){
  $raw=([string]$jobName+'|'+[string]$rp.part+'|'+[string]$rp.sourcePath)
  try{
    $sha=[Security.Cryptography.SHA256]::Create()
    $bytes=[Text.Encoding]::UTF8.GetBytes($raw)
    $hash=$sha.ComputeHash($bytes)
    return ([BitConverter]::ToString($hash)-replace '-','').Substring(0,16)
  }catch{
    $fallback=(($raw-replace '[^A-Za-z0-9]+','_').Trim('_'))
    if([string]::IsNullOrWhiteSpace($fallback)){return 'CLLINK'}
    return $fallback.Substring(0,[math]::Min(40,$fallback.Length))
  }
}
function SN-Apply-WorkspacePartData($app,$requestParts,[string]$jobName='',[string]$linkFile=''){
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
    $usedIndices += (SN-Scalar-Int -value $found.index -default -1)
    $partObj=$found.part
    $qty=SN-Scalar-Int -value $rp.qty -default 1
    if($qty -lt 1){$qty=1}
    $material=[string]$rp.sigmaMaterial
    $thickness=SN-Scalar-Number -value $rp.thicknessMm -default ([double]::NaN)
    $row=[ordered]@{
      part=$targetName
      index=(SN-Scalar-Int -value $found.index -default -1)
      qty=$qty
      material=$material
      thickness=$thickness
      quantityProperty=''
      materialProperty=''
      thicknessProperty=''
      warnings=@()
    }

    if(-not [string]::IsNullOrWhiteSpace($material)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @('Material','MaterialName','PartMaterial','MaterialDescription','MaterialType','Mat','MatName') -value $material -expectedText $material
      if(-not $setInfo){
        throw ('CL material "'+$material+'" could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.materialProperty=[string]$setInfo.path
    }

    if(-not [double]::IsNaN($thickness)){
      $setInfo=SN-Set-PartField -partObj $partObj -names @('Thickness','SheetThickness','MaterialThickness','ThicknessValue','PartThickness','Thk','Thick') -value $thickness -expectedNumber $thickness
      if(-not $setInfo){
        throw ('CL thickness '+$thickness+'mm could not be written and verified on SigmaNEST part "'+$targetName+'".')
      }
      $row.thicknessProperty=[string]$setInfo.path
    }

    # SigmaNEST's Part Parameters dialog calls the production quantity "Number To Nest".
    # Prefer that canonical field and other nest-quantity aliases. BatchQty is last because
    # SigmaNEST uses it in task/batch context and it must not override the part quantity.
    $setQty=SN-Set-QtyToNest -partObj $partObj -qty $qty
    $setInfo=if($setQty){[pscustomobject]@{path=$setQty;object=$partObj}}else{SN-Set-PartField -partObj $partObj -names @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity','BatchQuantity','BatchQty') -value $qty -expectedInt $qty}
    if(-not $setInfo){
      throw ('CL quantity '+$qty+' could not be written and verified on SigmaNEST part "'+$targetName+'".')
    }
    $row.quantityProperty=[string]$setInfo.path
    $drawingSet=SN-Set-PartField -partObj $partObj -names @('DrawingNumber','DrawingNo','DwgNumber','DwgNo','PartNumber','PartNo') -value $targetName -expectedText $targetName
    $workSet=SN-Set-PartField -partObj $partObj -names @('WONumber','WorkOrder','WorkOrderNumber','OrderNumber') -value $jobName -expectedText $jobName
    if($drawingSet){$row.drawingProperty=[string]$drawingSet.path}
    if($workSet){$row.workOrderProperty=[string]$workSet.path}
    $updated += [pscustomobject]$row
  }

  if(-not [string]::IsNullOrWhiteSpace($linkFile)){
    $records=@()
    foreach($rp in @($requestParts)){
      $m=@($updated|Where-Object {$_.part -eq [string]$rp.part}|Select-Object -First 1)
      $records += [pscustomobject]@{
        linkKey=$(if($m.Count){[string]$m[0].linkKey}else{SN-Make-CLLinkKey -jobName $jobName -rp $rp})
        jobName=[string]$jobName
        clPart=[string]$rp.part
        clMaterial=[string]$rp.material
        sigmaMaterial=[string]$rp.sigmaMaterial
        clThickness=[string]$rp.thicknessMm
        clQuantity=(SN-Scalar-Int -value $rp.qty -default 1)
        batchMultiplier=(SN-Scalar-Int -value $rp.batchMultiplier -default 1)
        sourceType=[string]$rp.sourceType
        sourcePath=[string]$rp.sourcePath
        sourceSheets=@($rp.sourceSheets)
        sourceRows=@($rp.sourceRows)
        sigmaNestPartIndex=$(if($m.Count){SN-Scalar-Int -value $m[0].index -default -1}else{-1})
      }
    }
    try{
      $parent=Split-Path -Parent $linkFile
      if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
      [ordered]@{schema='cl-ws-creator/cl-data-link/1.0';jobName=$jobName;createdUtc=(Get-Date).ToUniversalTime().ToString('o');records=$records} |
        ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $linkFile -Encoding UTF8
    }catch{}
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
    # Verify the actual Part Parameters quantity field, not a task-only batch field.
    $q=SN-Read-PartField -partObj $found.part -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $qty
    if($null -eq $q){throw ('Post-save verification failed for "'+$target+'": quantity is not '+$qty+'.')}
  }
  return $true
}
function SN-Parts-Count($app){try{return SN-Scalar-Int -value $app.PartsList.Count -default 0}catch{return 0}}

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
  if([string]::IsNullOrWhiteSpace($sourcePath)){throw 'SigmaNEST geometry sourcePath is required; refusing implicit PRS fallback.'}
  $sourceType=([string]$sourceType).Trim().ToUpperInvariant()
  $extension=[IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
  if($sourceType -eq 'DXF' -and $extension -ne '.dxf'){throw ('Geometry source type is DXF but the selected file is not a .DXF: '+$sourcePath)}
  if($sourceType -eq 'PRS' -and $extension -ne '.prs'){throw ('Geometry source type is PRS but the selected file is not a .PRS: '+$sourcePath)}
  if($sourceType -notin @('DXF','PRS')){throw ('Unsupported geometry source type "'+$sourceType+'". Expected DXF or PRS.')}
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $errors=@()

  if($sourceType -eq 'DXF'){
    # Use the exact .DXF path selected by the matcher. Do not probe the
    # PartsList.Import overloads: some SigmaNEST COM registrations bind the
    # reflection argument array incorrectly and surface "System.Object[] ->
    # System.Int32" conversion errors. LoadPart accepts the exact path and
    # keeps DXF selection deterministic; PRS fallback is forbidden here.
    try{
      [void]$app.LoadPart([string]$sourcePath)
      return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType='DXF';method='LoadPart-DXF-GEOMETRY'}
    }catch{
      $errors+=('LoadPart-DXF: '+(SN-ErrorText $_))
    }
    throw ('SigmaNEST could not load the selected .DXF "'+$label+'". No PRS fallback was attempted. '+($errors -join ' | '))
  }

  # PRS is an explicit fallback only when the matcher selected a PRS source.
  try{
    # LoadPart may queue the PRS geometry until CreatePartsListForNewPartsInWS;
    # the caller commits it and verifies the resulting PartsList entry.
    [void]$app.LoadPart([string]$sourcePath)
    return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType='PRS';method='LoadPart-PRS-GEOMETRY'}
  }catch{
    throw ('SigmaNEST could not load the selected .PRS "'+$label+'". '+(SN-ErrorText $_))
  }
}
function SN-Set-TaskMaterialAndThickness($app,$requestParts){
  $taskCount=0
  try{$taskCount=(SN-Scalar-Int -value $app.TasksList.Count -default 0)}catch{}
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
      try{$taskPartCount=(SN-Scalar-Int -value $task.PartsList.Count -default 0)}catch{}
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
  try{$taskCount=(SN-Scalar-Int -value $app.TasksList.Count -default 0)}catch{}
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
      try{$taskPartCount=(SN-Scalar-Int -value $task.PartsList.Count -default 0)}catch{}
      if($taskPartCount -le 0){continue}

      for($pi=0;$pi -lt $taskPartCount -and -not $matched;$pi++){
        $taskPart=$null
        try{$taskPart=$task.PartsList.Items($pi)}catch{continue}
        if($null -eq $taskPart){continue}

        $taskPartName=''
        try{$taskPartName=[string]$taskPart.Name}catch{}
        if([string]::IsNullOrWhiteSpace($taskPartName)){continue}
        if(-not $taskPartName.Equals($targetName,[StringComparison]::OrdinalIgnoreCase)){continue}

        $set=SN-Set-QtyToNest -partObj $taskPart -qty $quantity
        $readBack=$null
        # QtyToNest is the actual SigmaNEST COM field behind the Part Parameters
        # "Number To Nest" value. Only fall back to generic aliases if direct COM
        # access is unavailable.
        if(-not $set){ foreach($propertyName in @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity','BatchQuantity','BatchQty')){
          try{
            $taskPart.$propertyName=$quantity
            $readBack=SN-Scalar-Number -value $taskPart.$propertyName -default [double]::NaN
            if(-not [double]::IsNaN([double]$readBack) -and $readBack -eq $quantity){
              $set=$propertyName
              break
            }
          }catch{}
        } }

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
    foreach($x in @($Request.parts)){
      $source=[string]$x.sourcePath
      if([string]::IsNullOrWhiteSpace($source)){throw ('CL part "'+[string]$x.part+'" has no explicit geometry sourcePath. Refusing implicit PRS fallback.')}
      $sourceType=([string]$x.sourceType).Trim().ToUpperInvariant()
      if([string]::IsNullOrWhiteSpace($sourceType)){throw ('CL part "'+[string]$x.part+'" has no explicit geometry sourceType. Expected DXF or PRS.')}
      $load=SN-Queue-Geometry -app $app -sourcePath $source -sourceType $sourceType -clData ([pscustomobject]@{sigmaMaterial=[string]$x.sigmaMaterial;thicknessMm=$x.thicknessMm;qty=$x.qty})
      $qty=SN-Scalar-Int -value $x.qty -default 1
      if($qty -lt 1){$qty=1}
      $created += [pscustomobject]@{part=[string]$x.part;qty=$qty;material=[string]$x.sigmaMaterial;thickness=SN-Scalar-Number -value $x.thicknessMm -default ([double]::NaN);sourcePath=$source;sourceType=[string]$x.sourceType;importMethod=[string]$load.method;matchType=[string]$x.matchType;batchMultiplier=SN-Scalar-Int -value $x.batchMultiplier -default 1}
      $queued += [pscustomobject]@{part=[string]$x.part;qty=$qty;sigmaMaterial=[string]$x.sigmaMaterial;thicknessMm=$x.thicknessMm;sourcePath=$source;sourceType=[string]$x.sourceType;batchMultiplier=SN-Scalar-Int -value $x.batchMultiplier -default 1;sourceSheets=@($x.sourceSheets);sourceRows=@($x.sourceRows)}
    }
    if($queued.Count -eq 0){throw 'No geometry was found to import.'}
    $phase='COMMIT_IMPORTED_PARTS'
    [void]$app.CreatePartsListForNewPartsInWS()
    $after=SN-Parts-Count $app
    if($after -lt ($before+$queued.Count)){throw ('SigmaNEST committed '+($after-$before)+' part(s) but '+$queued.Count+' were requested.')}
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $queued -jobName $safe -linkFile ([string]$Request.clLinkFile)

    # IMPORT_ONLY intentionally stops here. Do not create TasksList entries and
    # do not call AutoTask. The operator must be able to inspect/correct the
    # imported PartsList data before any task is generated.
    $phase='SAVE_WORKSPACE'
    [void](SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'CL-data geometry import')
    # Do not reload the saved workspace through SNApp here. SigmaNEST.SNApp
    # attaches to the active SigmaNEST UI, and a background LoadWorkSpaceFile
    # can disturb the operator's active workspace/Part Parameters dialog.
    # The imported PartsList was already verified before save; after save we only
    # verify that the .ws file exists and is non-empty.
    $phase='VERIFY_SAVED_CL_PART_DATA'
    if(-not (Test-Path -LiteralPath $wsPath)){
      throw ('SigmaNEST reported a successful save but the workspace file was not created: '+$wsPath)
    }
    try{
      if((Get-Item -LiteralPath $wsPath).Length -le 0){
        throw ('SigmaNEST created an empty workspace file: '+$wsPath)
      }
    }catch{throw ('Saved workspace verification failed: '+$_.Exception.Message)}
    # Verification returns a boolean. Suppress it so this function emits exactly
    # one result object; otherwise PowerShell combines the boolean and result
    # object into System.Object[], which breaks the worker's Int32 conversions.
    [void](SN-Verify-WorkspaceCLData -app $app -requestParts $queued)
    return [pscustomobject]@{
      ok=$true
      phase='IMPORT_COMPLETE'
      wsPath=$wsPath
      parts=$created
      partCount=$created.Count
      partUpdates=@($partUpdates)
      clLinkFile=[string]$Request.clLinkFile
      tasksCreated=0
      message=('Geometry imported and CL material/thickness/quantity applied to SigmaNEST PartsList, verified, and saved to '+$wsPath+'. Tasks/AutoTask were not run; review the workspace before using AutoTask.')
      sourceCount=$created.Count
    }
  }catch{
    $exists=$false
    try{$exists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false
      phase=$phase
      wsPath=$(if($exists){[string]$wsPath}else{''})
      parts=$created
      partCount=@($created).Count
      partUpdates=@($partUpdates)
      tasksCreated=0
      checkpointSaved=$exists
      error=$_.Exception.Message
      message=$(if($exists){'SigmaNEST created a workspace checkpoint before the failure: '+$wsPath}else{'Geometry import failed before a verified workspace save.'})
    }
  }finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}
function SN-Set-TaskNameAndBatch($app,$requestParts){
  $taskCount=0;try{$taskCount=SN-Scalar-Int -value $app.TasksList.Count -default 0}catch{}
  $map=@{};foreach($rp in @($requestParts)){$map[[string]$rp.part]=$rp}
  $results=@();$warnings=@()
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$app.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $names=@();$materials=@();$thks=@();$multis=@()
    $pc=0;try{$pc=(SN-Scalar-Int -value $task.PartsList.Count -default 0)}catch{}
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
    [void]$app.LoadWorkSpaceFile([string]$wsPath)
    $phase='APPLY_CL_PART_DATA'
    $partUpdates=SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts -jobName ([string]$Request.jobName)

    # An imported workspace contains PartsList entries but not necessarily the
    # task objects that SigmaNEST AutoTask expects. Build the TasksList from the
    # imported workspace first, then let AutoTask organize/nest those tasks.
    $phase='CREATE_TASKS_FOR_IMPORTED_PARTS'
    [void]$app.CreateTasksListForNewPartsInWS()
    Start-Sleep -Milliseconds 500
    $taskCountBefore=0
    try{$taskCountBefore=SN-Scalar-Int -value $app.TasksList.Count -default 0}catch{}
    if($taskCountBefore -le 0){
      throw 'SigmaNEST could not create TasksList entries from the imported workspace parts.'
    }
    # Force the visible SigmaNEST UI to rebuild its tree before AutoTask.
    try{[void]$app.RefreshTreeView()}catch{}
    try{[void]$app.Redraw()}catch{}

    $phase='AUTO_TASK'
    [void]$app.AutoTask()
    Start-Sleep -Milliseconds 1500
    $taskCount=0
    try{$taskCount=SN-Scalar-Int -value $app.TasksList.Count -default 0}catch{}
    if($taskCount -le 0){
      throw 'SigmaNEST AutoTask completed but no TasksList entries are present.'
    }
    $phase='APPLY_TASK_CL_DATA'
    SN-Set-TaskMaterialAndThickness -app $app -requestParts $Request.parts
    SN-Set-TaskPartQuantity -app $app -requestParts $Request.parts
    $phase='LABEL_AND_BATCH'
    $taskData=SN-Set-TaskNameAndBatch -app $app -requestParts $Request.parts
    $phase='SAVE'
    $save=SN-Save-WorkspaceVerified -app $app -wsPath $wsPath -label 'AutoTask'
    # Refresh the existing SigmaNEST UI without reloading the workspace. Reloading
    # here can invalidate Part Parameters windows; tree refresh is sufficient to
    # expose the newly-created tasks.
    try{[void]$app.RefreshTreeView()}catch{}
    try{[void]$app.Redraw()}catch{}
    $ok=($taskData.warnings.Count -eq 0)
    return [pscustomobject]@{
      ok=$ok;phase='AUTOTASK_COMPLETE';wsPath=$wsPath;tasksCreated=$taskCount
      partCount=(SN-Parts-Count $app);importedCount=(SN-Parts-Count $app)
      parts=@($Request.parts)
      taskData=$taskData.tasks;warnings=$taskData.warnings;partUpdates=@($partUpdates)
      message=$(if($ok){'AutoTask created, labeled and batched '+$taskCount+' task(s).'}else{'AutoTask created '+$taskCount+' task(s) with warnings; see the Release Summary.'})
    }
  }catch{
    return [pscustomobject]@{
      ok=$false;phase=$phase;wsPath=$wsPath;tasksCreated=0
      partCount=$(try{SN-Parts-Count $app}catch{0})
      importedCount=$(try{SN-Parts-Count $app}catch{0})
      taskData=@();warnings=@($_.Exception.Message)
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
      $sourcePath=[string]$x.sourcePath;$sourceType=([string]$x.sourceType).Trim().ToUpperInvariant()
      if([string]::IsNullOrWhiteSpace($sourcePath)){throw ('CL part "'+[string]$x.part+'" has no explicit geometry sourcePath. Refusing implicit PRS fallback.')}
      if([string]::IsNullOrWhiteSpace($sourceType)){throw ('CL part "'+[string]$x.part+'" has no explicit geometry sourceType. Expected DXF or PRS.')}

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
        importMethod=[string]$load.method
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

    $phase='COMMIT_IMPORTED_PARTS';[void]$app.CreatePartsListForNewPartsInWS()
    $afterParts=SN-Parts-Count $app
    if($afterParts -lt ($beforeParts+$queued.Count)){
      throw ('SigmaNEST committed '+($afterParts-$beforeParts)+' part(s) but '+$queued.Count+' part(s) were requested for import.')
    }

    # Autosave the geometry workspace immediately after the imported parts are
    # committed. This guarantees a usable .ws exists even if later Task Setup
    # automation fails or disconnects a COM proxy.
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
      throw ('SigmaNEST imported '+($afterParts-$beforeParts)+' part(s), but the automatic .ws save did not produce a file: '+$wsPath+'. '+($saveErrors -join ' | '))
    }

    $phase='CREATE_TASKS';[void]$app.CreateTasksListForNewPartsInWS()
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
    $phase='SAVE_WS';$app.SaveWorkSpaceFile([string]$wsPath)
    try{[void]$app.LoadWorkSpaceFile([string]$wsPath)}catch{};try{[void]$app.RefreshTreeView()}catch{};try{[void]$app.Redraw()}catch{}
    return [pscustomobject]@{ok=$true;creatorVersion='DIRECT-COM-2.14.0';phase='COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count;message=('SigmaNEST WS created: '+$wsPath)}
  }catch{
    $checkpointExists=$false
    try{$checkpointExists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false
      creatorVersion='DIRECT-COM-2.14.0'
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