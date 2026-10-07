param([Parameter(Mandatory=$true)][string]$RequestFile,[string]$ResultFile)
$ErrorActionPreference='Stop'
$CREATOR_VERSION='3.5.0'

function Write-Diagnostic([string]$message,[string]$phase='STARTUP',[int]$exitCode=1){
  if(-not [string]::IsNullOrWhiteSpace($ResultFile)){
    try{
      $parent=Split-Path -Parent $ResultFile
      if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
      $obj=[ordered]@{ok=$false;creatorVersion=$CREATOR_VERSION;phase=$phase;exitCode=$exitCode;error=$message;requestFile=$RequestFile}
      [IO.File]::WriteAllText($ResultFile,($obj|ConvertTo-Json -Depth 8 -Compress),(New-Object System.Text.UTF8Encoding($false)))
    }catch{}
  }
}

Write-Diagnostic 'SigmaNEST creator process started.' 'STARTUP' 0

function Out($o){
  $json=$o|ConvertTo-Json -Depth 16 -Compress
  if(-not [string]::IsNullOrWhiteSpace($ResultFile)){
    try{
      $parent=Split-Path -Parent $ResultFile
      if($parent){New-Item -ItemType Directory -Path $parent -Force|Out-Null}
      [IO.File]::WriteAllText($ResultFile,$json,(New-Object System.Text.UTF8Encoding($false)))
    }catch{}
  }
  Write-Output $json
}

function Invoke-ComMethod($obj,[string]$name,[object[]]$args=@()){
  $obj.GetType().InvokeMember(
    $name,
    [Reflection.BindingFlags]::InvokeMethod,
    $null,
    $obj,
    $args
  )
}

function Try-Set($obj,[string[]]$names,$value){
  if($null -eq $obj){return $null}
  foreach($name in $names){
    if([string]::IsNullOrWhiteSpace($name)){continue}
    try{$obj.$name=$value;return $name}catch{}
    try{
      $obj.GetType().InvokeMember(
        $name,
        [Reflection.BindingFlags]::SetProperty,
        $null,
        $obj,
        @($value)
      )|Out-Null
      return $name
    }catch{}
  }
  return $null
}

function Set-Required($obj,[string[]]$names,$value,[string]$label){
  $set=Try-Set -obj $obj -names $names -value $value
  if(-not $set){
    throw ('SigmaNEST part does not expose a writable '+$label+' property. Tried: '+($names -join ', '))
  }
  return $set
}

function Safe-Set($obj,[string]$name,$value){
  [void](Try-Set -obj $obj -names @($name) -value $value)
}

function Get-PartsCount($app){
  try{return [int]$app.PartsList.Count}catch{return 0}
}

function Get-NewPart($app,[int]$beforeCount,[string]$label){
  $afterCount=Get-PartsCount $app
  if($afterCount -le $beforeCount){
    throw ('SigmaNEST loaded "'+$label+'" but did not add it to the workspace PartsList.')
  }
  # The X1.4 PartsList interface did not expose an Items() member in the
  # verified COM metadata. Prefer COM enumeration, then try common indexers.
  try{
    $last=$null
    foreach($item in $app.PartsList){$last=$item}
    if($null -ne $last){return $last}
  }catch{}
  foreach($member in @('Item','Items','get_Item')){
    try{
      $obj=$app.PartsList.GetType().InvokeMember(
        $member,
        [Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::GetProperty,
        $null,$app.PartsList,@([int]($afterCount-1))
      )
      if($null -ne $obj){return $obj}
    }catch{}
  }
  throw ('SigmaNEST added "'+$label+'" but the new PartsList item could not be accessed. PartsList.Count='+$afterCount)
}

function Add-GeometryToWorkspace($app,[string]$sourcePath,[string]$sourceType){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $sourceType=([string]$sourceType).Trim().ToUpperInvariant()
  $extension=[IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
  if($sourceType -eq 'DXF' -and $extension -ne '.dxf'){throw ('DXF source type requires a .DXF path: '+$sourcePath)}
  if($sourceType -eq 'PRS' -and $extension -ne '.prs'){throw ('PRS source type requires a .PRS path: '+$sourcePath)}
  if($sourceType -notin @('DXF','PRS')){throw ('Unsupported geometry source type "'+$sourceType+'".')}
  $label=[IO.Path]::GetFileName($sourcePath)
  $before=Get-PartsCount $app
  $errors=@()

  if($sourceType -eq 'DXF'){
    try{
      $settings=$null
      try{$settings=New-Object -ComObject SigmaNEST.SNPartImportSettings}catch{}
      foreach($importId in @(0,1,2,3)){
        $beforeTry=Get-PartsCount $app
        try{
          $result=Invoke-ComMethod $app.PartsList 'Import' @([string]$sourcePath,[int]$importId,$settings)
          Start-Sleep -Milliseconds 150
          try{Invoke-ComMethod $app 'CreatePartsListForNewPartsInWS' @()|Out-Null}catch{}
          $afterTry=Get-PartsCount $app
          if($afterTry -gt $beforeTry -or ($null -ne $result -and [bool]$result)){
            return Get-NewPart $app $before $label
          }
        }catch{
          $errors += ('PartsList.Import id '+$importId+': '+$_.Exception.Message)
        }
      }
    }catch{
      $errors += ('DXF import setup: '+$_.Exception.Message)
    }
    try{
      $beforeLoad=Get-PartsCount $app
      [void](Invoke-ComMethod $app 'LoadPart' @([string]$sourcePath))
      $afterLoad=Get-PartsCount $app
      if($afterLoad -gt $beforeLoad){return Get-NewPart $app $before $label}
    }catch{$errors += ('LoadPart-DXF: '+$_.Exception.Message)}
    throw ('SigmaNEST could not import selected .DXF "'+$label+'". No PRS fallback was attempted. '+($errors -join ' | '))
  }

  try{
    [void](Invoke-ComMethod $app 'LoadPart' @([string]$sourcePath))
    Start-Sleep -Milliseconds 100
    return Get-NewPart $app $before $label
  }catch{
    $errors += ('LoadPart-PRS: '+$_.Exception.Message)
  }
  throw ('SigmaNEST could not import selected .PRS "'+$label+'". '+($errors -join ' | '))
}
try{
  $phase='READ_REQUEST'
  $req=Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8|ConvertFrom-Json
  $phase='CREATE_COM'
  $app=New-Object -ComObject SigmaNEST.SNApp
  $auto=New-Object -ComObject SigmaNEST.SNAutomation

  $phase='RESOLVE_WS_PATH'
  $wsDir=[string]$req.wsDirectory
  if([string]::IsNullOrWhiteSpace($wsDir)){
    $wsRoot=[string]$req.wsRoot
    if([string]::IsNullOrWhiteSpace($wsRoot)){
      $libRoot=[string]$req.libraryRoot
      if(-not [string]::IsNullOrWhiteSpace($libRoot)){
        try{
          $parent=[IO.Directory]::GetParent($libRoot)
          if($parent){$wsRoot=$parent.FullName}
        }catch{}
      }
    }
    if(-not [string]::IsNullOrWhiteSpace($wsRoot)){
      try{$wsDir=Join-Path -Path $wsRoot -ChildPath 'WS'}catch{}
    }
  }
  if([string]::IsNullOrWhiteSpace($wsDir)){
    throw 'SigmaNEST WS output folder could not be determined. Configure the SigmaNEST WS folder (PathID 0) or verify the PRS library path.'
  }
  $wsDir=[IO.Path]::GetFullPath($wsDir.Trim())
  if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}

  $job=[string]$req.jobName
  if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
  $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim()
  if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}

  $wsPath=Join-Path -Path $wsDir -ChildPath ($safe+'.ws')
  if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}

  $phase='SIGMANEST_WORKSPACE_INIT'
  # SNApp creates an automation workspace when instantiated. Do not call
  # SNAutomation.FileNew(): that method is not exposed by the verified X1.4
  # SNAutomation interface and was a primary source of creator failures.
  try{
    $members=@($auto.GetType().GetMembers() | ForEach-Object { $_.Name })
    if($members -contains 'FileNew'){ Invoke-ComMethod $auto 'FileNew' @() | Out-Null }
  }catch{
    # FileNew is optional. A newly-created SNApp is already a clean workspace.
  }
  try{$app.PartsLibrary.Directory=[string]$req.libraryRoot}catch{}

  $created=@()
  foreach($x in @($req.parts)){
    $phase='IMPORT_PART'
    $sourcePath=[string]$x.sourcePath
    $sourceType=[string]$x.sourceType
    if([string]::IsNullOrWhiteSpace($sourcePath)){
      throw ('Request part "'+[string]$x.part+'" has no explicit sourcePath. Refusing implicit PRS fallback.')
    }
    $sourceType=([string]$sourceType).Trim().ToUpperInvariant()
    if([string]::IsNullOrWhiteSpace($sourceType)){
      throw ('Request part "'+[string]$x.part+'" has no explicit sourceType. Expected DXF or PRS.')
    }

    $part=Add-GeometryToWorkspace $app $sourcePath $sourceType

    # BatchQty is the verified quantity field. Material is required for a new
    # DXF and is also explicitly set for PRS parts so Auto Task can classify them.
    $quantity=[int][math]::Round([double]$x.qty)
    if($quantity -lt 1){$quantity=1}
    $quantityProperty=Set-Required -obj $part -names @('BatchQty','QtyToNest','Quantity','Qty') -value $quantity -label 'quantity'
    $materialProperty=$null
    if(-not [string]::IsNullOrWhiteSpace([string]$x.sigmaMaterial)){
      $materialProperty=Set-Required -obj $part -names @('Material') -value ([string]$x.sigmaMaterial) -label 'material'
    }
    $thicknessProperty=$null
    if($x.thicknessMm -ne $null -and -not [double]::IsNaN([double]$x.thicknessMm)){
      # Thickness is not guaranteed to be writable on every X1.4 Part COM
      # object. Try the known names, but do not reject an otherwise valid PRS.
      $thicknessProperty=Try-Set -obj $part -names @('Thickness','SheetThickness','Thk') -value ([double]$x.thicknessMm)
    }
    $sourceProperty=Try-Set -obj $part -names @('SourceFilePath','SourcePath') -value ([string]$sourcePath)
    Safe-Set $part 'WONumber' $safe
    Safe-Set $part 'DrawingNumber' ([string]$x.part)

    $created += [pscustomobject]@{
      part=$(try{[string]$part.Name}catch{[string]$x.part})
      qty=$quantity
      material=$(try{[string]$part.Material}catch{[string]$x.sigmaMaterial})
      thickness=$(try{[string]$part.Thickness}catch{[string]$x.thicknessMm})
      quantityProperty=$quantityProperty
      materialProperty=$materialProperty
      thicknessProperty=$thicknessProperty
      sourcePath=$sourcePath
      sourceType=$sourceType
      sourceProperty=$sourceProperty
      batchMultiplier=$x.batchMultiplier
      taskBatches=@($x.taskBatches)
    }
  }

  try{Invoke-ComMethod $app 'CreateTasksListForNewPartsInWS' @()|Out-Null}catch{}
  $phase='SAVE_WS'
  Invoke-ComMethod $app 'SaveWorkSpaceFile' @([string]$wsPath) | Out-Null
  try{$app.LoadWorkSpaceFile($wsPath)}catch{}
  try{$app.RefreshTreeView()}catch{}
  try{$app.Redraw()}catch{}

  [pscustomobject]@{
    ok=$true
    creatorVersion=$CREATOR_VERSION
    wsPath=$wsPath
    parts=$created
    partCount=$created.Count
    message=('SigmaNEST WS created: '+$wsPath)
  }|Out
}catch{
  $err=$_.Exception.Message
  Write-Diagnostic $err $phase 1
  [pscustomobject]@{
    ok=$false
    creatorVersion=$CREATOR_VERSION
    error=$_.Exception.Message
    category=$_.CategoryInfo.ToString()
  }|Out
  return
}
