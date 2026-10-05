# Pure SigmaNEST COM function library. Loaded by server.ps1, which runs in STA.
# No top-level execution and no exit statements.

function SN-Invoke-ComMethod($obj,[string]$name,[object[]]$args=@()){
  # Use reflection only for calls whose COM signature is not fixed in our
  # verified SigmaNEST interface. Verified SNApp methods are called directly
  # elsewhere below so PowerShell's COM binder supplies the correct signature.
  $obj.GetType().InvokeMember($name,[Reflection.BindingFlags]::InvokeMethod,$null,$obj,$args)
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

function SN-Add-Geometry($app,[string]$sourcePath,[string]$sourceType){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $before=SN-Parts-Count $app
  $errors=@()
  try{
    # Verified SigmaNEST X1.4 interface: SNApp.LoadPart(string).
    # Do not use Type.InvokeMember here; it binds the COM method incorrectly
    # on this installation and reports a parameter-count mismatch.
    $loaded=$app.LoadPart([string]$sourcePath)
    if([bool]$loaded){
      try{$app.CreatePartsListForNewPartsInWS()}catch{}
      return SN-Get-NewPart $app $before $label
    }
    $errors+='LoadPart returned False'
  }catch{$errors+=('LoadPart: '+$_.Exception.Message)}
  if($sourceType -eq 'DXF'){
    try{
      $settings=$null;try{$settings=New-Object -ComObject SigmaNEST.SNPartImportSettings}catch{}
      foreach($importId in @(0,1,2,3)){
        $beforeTry=SN-Parts-Count $app
        try{
          $result=SN-Invoke-ComMethod $app.PartsList 'Import' @([string]$sourcePath,[int]$importId,$settings)
          Start-Sleep -Milliseconds 150
          try{SN-Invoke-ComMethod $app 'CreatePartsListForNewPartsInWS' @()|Out-Null}catch{}
          $afterTry=SN-Parts-Count $app
          if($afterTry -gt $beforeTry){return SN-Get-NewPart $app $before $label}
          if($result -ne $null -and [bool]$result){$afterTry=SN-Parts-Count $app;if($afterTry -gt $beforeTry){return SN-Get-NewPart $app $before $label}}
        }catch{$errors+=('PartsList.Import id '+$importId+': '+$_.Exception.Message)}
      }
    }catch{$errors+=('DXF import fallback: '+$_.Exception.Message)}
  }
  throw ('SigmaNEST could not import "'+$label+'". '+($errors -join ' | '))
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
    $phase='IMPORT_PARTS';$created=@()
    foreach($x in @($Request.parts)){
      $sourcePath=[string]$x.sourcePath;$sourceType=[string]$x.sourceType
      if([string]::IsNullOrWhiteSpace($sourcePath)){$sourcePath=[string]$x.prsPath;if(-not $sourceType){$sourceType='PRS'}}
      if([string]::IsNullOrWhiteSpace($sourcePath)){continue}
      $part=SN-Add-Geometry $app $sourcePath $sourceType
      $quantity=[int][math]::Round([double]$x.qty);if($quantity -lt 1){$quantity=1}
      $quantityProperty=SN-Set-Required -obj $part -names @('BatchQty','QtyToNest','Quantity','Qty') -value $quantity -label 'quantity'
      $materialProperty=$null;if(-not [string]::IsNullOrWhiteSpace([string]$x.sigmaMaterial)){$materialProperty=SN-Set-Required -obj $part -names @('Material') -value ([string]$x.sigmaMaterial) -label 'material'}
      $thicknessProperty=$null;if($x.thicknessMm -ne $null -and -not [double]::IsNaN([double]$x.thicknessMm)){$thicknessProperty=SN-Try-Set -obj $part -names @('Thickness','SheetThickness','Thk') -value ([double]$x.thicknessMm)}
      $sourceProperty=SN-Try-Set -obj $part -names @('SourceFilePath','SourcePath') -value ([string]$sourcePath);SN-Safe-Set $part 'WONumber' $safe;SN-Safe-Set $part 'DrawingNumber' ([string]$x.part)
      $created += [pscustomobject]@{part=$(try{[string]$part.Name}catch{[string]$x.part});qty=$quantity;material=$(try{[string]$part.Material}catch{[string]$x.sigmaMaterial});thickness=$(try{[string]$part.Thickness}catch{[string]$x.thicknessMm});quantityProperty=$quantityProperty;materialProperty=$materialProperty;thicknessProperty=$thicknessProperty;sourcePath=$sourcePath;sourceType=$sourceType;sourceProperty=$sourceProperty;batchMultiplier=$x.batchMultiplier;taskBatches=@($x.taskBatches)}
    }
    $phase='CREATE_TASKS';try{$app.CreateTasksListForNewPartsInWS()}catch{}
    $phase='SAVE_WS';$app.SaveWorkSpaceFile([string]$wsPath)
    try{$app.LoadWorkSpaceFile([string]$wsPath)}catch{};try{$app.RefreshTreeView()}catch{};try{$app.Redraw()}catch{}
    return [pscustomobject]@{ok=$true;creatorVersion='DIRECT-COM-1.0';phase='COMPLETE';wsPath=$wsPath;parts=$created;partCount=$created.Count;message=('SigmaNEST WS created: '+$wsPath)}
  }catch{return [pscustomobject]@{ok=$false;creatorVersion='DIRECT-COM-1.0';phase=$phase;error=$_.Exception.Message;category=$_.CategoryInfo.ToString()}}
  finally{if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}}
}