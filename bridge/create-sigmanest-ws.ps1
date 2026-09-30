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
  $auto.FileNew()
  try{$app.PartsLibrary.Directory=[string]$req.libraryRoot}catch{}
  $created=@()
  foreach($x in @($req.parts)){
    $prs=[string]$x.prsPath
    if([string]::IsNullOrWhiteSpace($prs)){continue}
    if(-not(Test-Path -LiteralPath $prs)){throw ('PRS not found: '+$prs)}
    $part=$null
    try{$part=$app.PartsList.AddfromLibrary($prs)}catch{
      $leaf=[IO.Path]::GetFileNameWithoutExtension($prs)
      $part=$app.PartsList.AddfromLibrary($leaf)
    }
    if($null -eq $part){throw ('SigmaNEST could not load part: '+$prs)}
    try{$part.QtyToNest=[int][math]::Round([double]$x.qty)}catch{}
    try{$part.BatchQty=[int][math]::Round([double]$x.qty)}catch{}
    try{$part.Material=[string]$x.sigmaMaterial}catch{}
    try{$part.Thickness=[double]$x.thicknessMm}catch{}
    try{$part.WONumber=$safe}catch{}
    $created += [pscustomobject]@{part=$part.Name;qty=$part.QtyToNest;material=$part.Material;thickness=$part.Thickness;path=$part.Path}
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
