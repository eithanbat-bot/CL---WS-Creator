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

function SN-Parts-Count($app){try{return [int]$app.PartsList.Count}catch{return 0}}

function SN-Get-NewPart($app,[int]$beforeCount,[string]$label){
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
    $quantity=[int][math]::Round([double]$rp.qty)
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
      $quantity=[int][math]::Round([double]$x.qty);if($quantity -lt 1){$quantity=1}
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
    $phase='CREATE_TASKS';$app.CreateTasksListForNewPartsInWS()
    # Persist the geometry/tasks checkpoint before touching mutable Task Setup
    # COM interfaces. If a later Task Setup call invalidates a COM proxy, the
    # usable geometry workspace is still saved on disk.
    $phase='SAVE_GEOMETRY_CHECKPOINT';$app.SaveWorkSpaceFile([string]$wsPath)
    $phase='APPLY_TASK_ATTRIBUTES';SN-Set-TaskMaterialAndThickness -app $app -requestParts $Request.parts
    $phase='APPLY_TASK_QUANTITIES';SN-Set-TaskPartQuantity -app $app -requestParts $Request.parts
    $phase='SAVE_WS';$app.SaveWorkSpaceFile([string]$wsPath)
    try{$app.LoadWorkSpaceFile([string]$wsPath)}catch{};try{$app.RefreshTreeView()}catch{};try{$app.Redraw()}catch{}
    return [pscustomobject]@{ok=$true;creatorVersion='DIRECT-COM-1.3';phase='COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count;message=('SigmaNEST WS created: '+$wsPath)}
  }catch{
    $checkpointExists=$false
    try{$checkpointExists=Test-Path -LiteralPath ([string]$wsPath)}catch{}
    return [pscustomobject]@{
      ok=$false
      creatorVersion='DIRECT-COM-1.3'
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