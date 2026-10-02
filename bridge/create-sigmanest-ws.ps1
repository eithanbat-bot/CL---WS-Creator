param([Parameter(Mandatory=$true)][string]$RequestFile)
$ErrorActionPreference='Stop'
$CREATOR_VERSION='3.2.1'

function Out($o){$o|ConvertTo-Json -Depth 16 -Compress}

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
    try{
      $obj.$name=$value
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
  try{
    return $app.PartsList.Items($afterCount-1)
  }catch{
    throw ('SigmaNEST added "'+$label+'" but the new PartsList item could not be accessed: '+$_.Exception.Message)
  }
}

function Add-GeometryToWorkspace($app,[string]$sourcePath,[string]$sourceType){
  if(-not(Test-Path -LiteralPath $sourcePath)){throw ('Geometry not found: '+$sourcePath)}
  $label=[IO.Path]::GetFileName($sourcePath)
  $before=Get-PartsCount $app
  $errors=@()

  # The X1.4 SNApp exposes LoadPart(path). Use it first for both .PRS and .DXF.
  # On installations where DXF must go through the explicit import interface,
  # fall back to PartsList.Import(path, ImportID, settings).
  try{
    $loaded=Invoke-ComMethod $app 'LoadPart' @([string]$sourcePath)
    if([bool]$loaded){
      try{Invoke-ComMethod $app 'CreatePartsListForNewPartsInWS' @()|Out-Null}catch{}
      return Get-NewPart $app $before $label
    }
    $errors += 'LoadPart returned False'
  }catch{
    $errors += ('LoadPart: '+$_.Exception.Message)
  }

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
          if($afterTry -gt $beforeTry){
            return Get-NewPart $app $before $label
          }
          if($result -ne $null -and [bool]$result){
            $afterTry=Get-PartsCount $app
            if($afterTry -gt $beforeTry){return Get-NewPart $app $before $label}
          }
        }catch{
          $errors += ('PartsList.Import id '+$importId+': '+$_.Exception.Message)
        }
      }
    }catch{
      $errors += ('DXF import fallback: '+$_.Exception.Message)
    }
  }

  throw ('SigmaNEST could not import "'+$label+'". '+($errors -join ' | '))
}

try{
  $req=Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8|ConvertFrom-Json
  $app=New-Object -ComObject SigmaNEST.SNApp
  $auto=New-Object -ComObject SigmaNEST.SNAutomation
  $paths=New-Object -ComObject SigmaNEST.SNPaths

  $wsDir=[string]$req.wsDirectory
  if([string]::IsNullOrWhiteSpace($wsDir)){
    try{$wsDir=[string]$paths.GetPath(0)}catch{$wsDir=''}
  }
  if([string]::IsNullOrWhiteSpace($wsDir)){
    $libRoot=[string]$req.libraryRoot
    if(-not [string]::IsNullOrWhiteSpace($libRoot)){
      try{
        $parent=[IO.Directory]::GetParent($libRoot)
        if($parent){$wsDir=Join-Path -Path $parent.FullName -ChildPath 'WS'}
      }catch{}
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

  $auto.FileNew()
  try{$app.PartsLibrary.Directory=[string]$req.libraryRoot}catch{}

  $created=@()
  foreach($x in @($req.parts)){
    $sourcePath=[string]$x.sourcePath
    $sourceType=[string]$x.sourceType
    if([string]::IsNullOrWhiteSpace($sourcePath)){
      $sourcePath=[string]$x.prsPath
      if(-not $sourceType){$sourceType='PRS'}
    }
    if([string]::IsNullOrWhiteSpace($sourcePath)){continue}

    $part=Add-GeometryToWorkspace $app $sourcePath $sourceType

    # Quantity is known from the SigmaNEST X1.4 part interface as BatchQty.
    # Do not silently ignore a failed assignment: an incorrect WS is worse than
    # stopping the build for review.
    $quantity=[int][math]::Round([double]$x.qty)
    $quantityProperty=Set-Required -obj $part -names @('BatchQty','QtyToNest','Quantity','Qty') -value $quantity -label 'quantity'
    $materialProperty=$null
    if($x.sigmaMaterial){
      $materialProperty=Set-Required -obj $part -names @('Material') -value ([string]$x.sigmaMaterial) -label 'material'
    }
    $thicknessProperty=$null
    if($x.thicknessMm -ne $null -and -not [double]::IsNaN([double]$x.thicknessMm)){
      $thicknessProperty=Set-Required -obj $part -names @('Thickness','SheetThickness','Thk') -value ([double]$x.thicknessMm) -label 'thickness'
    }
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
      batchMultiplier=$x.batchMultiplier
      taskBatches=@($x.taskBatches)
    }
  }

  try{Invoke-ComMethod $app 'CreateTasksListForNewPartsInWS' @()|Out-Null}catch{}
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
  [pscustomobject]@{
    ok=$false
    creatorVersion=$CREATOR_VERSION
    error=$_.Exception.Message
    category=$_.CategoryInfo.ToString()
  }|Out
  exit 1
}
