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

function SN-Queue-Geometry($app,[string]$sourcePath,[string]$sourceType){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $errors=@()
  try{
    # Load the geometry into SigmaNEST's pending/new-parts collection.
    # CreatePartsListForNewPartsInWS is intentionally NOT called here:
    # invoking it once per part is extremely expensive for large CLs.
    [void]$app.LoadPart([string]$sourcePath)
    return [pscustomobject]@{ok=$true;label=$label;sourcePath=$sourcePath;sourceType=$sourceType;method='LoadPart'}
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
      $load=SN-Queue-Geometry -app $app -sourcePath $source -sourceType ([string]$x.sourceType)
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
        thicknessMm=$x.thicknessMm;workspaceIndex=$workspaceIndex
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