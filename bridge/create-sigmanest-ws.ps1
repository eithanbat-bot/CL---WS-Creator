param([Parameter(Mandatory=$true)][string]$RequestFile)
$ErrorActionPreference='Stop'
function Out($o){$o|ConvertTo-Json -Depth 12 -Compress}
try{
  $req=Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8|ConvertFrom-Json
  $app=New-Object -ComObject SigmaNEST.SNApp
  $auto=New-Object -ComObject SigmaNEST.SNAutomation
  $paths=New-Object -ComObject SigmaNEST.SNPaths
  $wsDir=$req.wsDirectory
  if([string]::IsNullOrWhiteSpace([string]$wsDir)){ $wsDir=$paths.GetPath(0) }
  if(-not(Test-Path -LiteralPath $wsDir)){New-Item -ItemType Directory -Path $wsDir -Force|Out-Null}
  $job=[string]$req.jobName
  if([string]::IsNullOrWhiteSpace($job)){throw 'Job name is required.'}
  $safe=($job -replace '[^A-Za-z0-9._ -]','_').Trim()
  if([string]::IsNullOrWhiteSpace($safe)){$safe='CL_JOB'}
  $wsPath=Join-Path $wsDir ($safe+'.ws')
  if(Test-Path -LiteralPath $wsPath){throw ('SigmaNEST WS already exists: '+$wsPath)}
  $auto.FileNew()
  try{$app.PartsLibrary.Directory=[string]$req.libraryRoot}catch{}
  $created=@()
  foreach($x in @($req.parts)){
    $prs=[string]$x.prsPath
    if([string]::IsNullOrWhiteSpace($prs)){continue}
    if(-not(Test-Path -LiteralPath $prs)){throw ('PRS not found: '+$prs)}
    $part=$null
    $leaf=[IO.Path]::GetFileNameWithoutExtension($prs)

    # Do not use PartsList.AddfromLibrary here. On some SigmaNEST X1.4
    # installations PowerShell's COM binder treats that member as a command
    # parameter call and produces:
    # "A positional parameter cannot be found that accepts argument ...PRS".
    #
    # SNApp.LoadPart(path) is the COM entry point exposed by the installation
    # diagnostics and returns a real success/failure value. Once a part is
    # loaded, CreatePartsListForNewPartsInWS() commits the newly loaded part
    # into the workspace PartsList.
    try{
      $loaded=$app.LoadPart([string]$prs)
    }catch{
      throw ('SigmaNEST LoadPart failed for "'+$leaf+'.PRS": '+$_.Exception.Message)
    }
    if(-not [bool]$loaded){
      throw ('SigmaNEST LoadPart returned False for "'+$leaf+'.PRS".')
    }

    $beforeCount=0
    try{$beforeCount=[int]$app.PartsList.Count}catch{}
    try{$app.CreatePartsListForNewPartsInWS()}catch{
      throw ('SigmaNEST could not add "'+$leaf+'.PRS" to the workspace PartsList: '+$_.Exception.Message)
    }

    $afterCount=0
    try{$afterCount=[int]$app.PartsList.Count}catch{}
    if($afterCount -le $beforeCount){
      throw ('SigmaNEST loaded "'+$leaf+'.PRS" but did not add it to the workspace PartsList.')
    }

    # Use the newly-created last PartsList item.
    try{$part=$app.PartsList.Items($afterCount-1)}catch{
      throw ('SigmaNEST added "'+$leaf+'.PRS" but its new PartsList item could not be accessed: '+$_.Exception.Message)
    }
    try{$part.QtyToNest=[int][math]::Round([double]$x.qty)}catch{}
    try{$part.Material=[string]$x.sigmaMaterial}catch{}
    try{$part.Thickness=[double]$x.thicknessMm}catch{}
    try{$part.WONumber=$safe}catch{}
    try{$part.DrawingNumber=[string]$x.part}catch{}
    $created += [pscustomobject]@{part=$part.Name;qty=$part.QtyToNest;material=$part.Material;thickness=$part.Thickness;path=$part.Path;batchMultiplier=$x.batchMultiplier}
  }
  $app.SaveWorkSpaceFile($wsPath)
  try{$app.LoadWorkSpaceFile($wsPath)}catch{}
  try{$app.RefreshTreeView()}catch{}
  try{$app.Redraw()}catch{}
  [pscustomobject]@{ok=$true;wsPath=$wsPath;parts=$created;partCount=$created.Count;message=('SigmaNEST WS created: '+$wsPath)}|Out
}catch{
  [pscustomobject]@{ok=$false;error=$_.Exception.Message;category=$_.CategoryInfo.ToString()}|Out
  exit 1
}
