param(
  [string]$DxfPath='Y:\SOLIDWORKS\UNIVERSAL COMPONENTS\PC-2A.DXF',
  [string]$PrsPath='S:\SNDataX1\PARTS\HSWU2 100.PRS'
)
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
. (Join-Path $Root 'bridge\sigmanest-com.ps1')
$outDir=Join-Path $env:TEMP 'CL-WS-Creator-MixedRealTest'
New-Item -ItemType Directory -Path $outDir -Force|Out-Null
$job='CLWS-MIX-'+[Guid]::NewGuid().ToString('N').Substring(0,8)
$wsPath=Join-Path $outDir ($job+'.ws')
$parts=@(
  [pscustomobject]@{part='PC-2A';qty=4;sigmaMaterial='MS';thicknessMm=4;sourcePath=$DxfPath;sourceType='DXF';matchType='SELFTEST';batchMultiplier=4;taskBatches=@();sourceSheets=@('DXF');sourceRows=@()},
  [pscustomobject]@{part='HSWU2 100';qty=2;sigmaMaterial='MS';thicknessMm=8;sourcePath=$PrsPath;sourceType='PRS';matchType='SELFTEST';batchMultiplier=4;taskBatches=@();sourceSheets=@('PRS');sourceRows=@()}
)
$importResult=$null;$autoResult=$null;$verifyAuto=$null;$verify=$null
try{
  foreach($source in @($DxfPath,$PrsPath)){if(-not(Test-Path -LiteralPath $source)){throw ('Self-test geometry fixture is missing: '+$source)}}
  $request=[pscustomobject]@{jobName=$job;wsDirectory=$outDir;libraryRoot='S:\SNDataX1\PARTS';clLinkFile='';parts=$parts}
  $importResult=@(Invoke-SigmaNestImportGeometry -Request $request)[-1]
  if(-not $importResult.ok){throw ('Mixed import failed at '+$importResult.phase+': '+$importResult.error)}
  if((SN-Scalar-Int -value $importResult.partCount -default 0) -ne 2){throw ('Mixed import expected 2 parts, got '+$importResult.partCount)}
  $autoRequest=[pscustomobject]@{jobName=$job;wsPath=$wsPath;parts=$parts}
  $autoResult=@(Invoke-SigmaNestAutoTask -Request $autoRequest)[-1]
  if(-not $autoResult.ok){throw ('Mixed AutoTask failed at '+$autoResult.phase+': '+$autoResult.message+'; '+(@($autoResult.warnings)-join ' | '))}
  if(@($autoResult.warnings|Where-Object{-not [string]::IsNullOrWhiteSpace([string]$_)}).Count -gt 0){throw ('Mixed AutoTask warnings: '+(@($autoResult.warnings)-join ' | '))}
  $verifyAuto=New-Object -ComObject SigmaNEST.SNAutomation
  [void]$verifyAuto.ResetSigmaNEST();[void]$verifyAuto.FileNew()
  $verify=New-Object -ComObject SigmaNEST.SNApp
  [void]$verify.LoadWorkSpaceFile($wsPath);Start-Sleep -Milliseconds 500
  $partCount=SN-Parts-Count $verify
  $taskCount=SN-Scalar-Int -value $verify.TasksList.Count -default 0
  if($partCount -ne 2){throw ('Mixed final workspace expected 2 parts, got '+$partCount)}
  if($taskCount -lt 2){throw ('Mixed final workspace should create separate 4mm and 8mm tasks; found '+$taskCount)}
  $taskData=@($autoResult.taskData)
  if($taskData.Count -ne 2){throw ('Expected two separate material/thickness task records; found '+$taskData.Count)}
  $expectedLabels=@('01 | MS | 4mm','02 | MS | 8mm')
  for($ti=0;$ti -lt 2;$ti++){
    if([string]$taskData[$ti].taskName -ne $expectedLabels[$ti]){throw ('Unexpected mixed task label at '+($ti+1)+': '+[string]$taskData[$ti].taskName)}
    if(-not [bool]$taskData[$ti].labelApplied){throw ('Task label write was not verified for '+$expectedLabels[$ti])}
    if(-not [bool]$taskData[$ti].batchApplied){throw ('Task batch multiplier write was not verified for '+$expectedLabels[$ti])}
    if((SN-Scalar-Int -value $taskData[$ti].batchMultiplier -default 0) -ne 4){throw ('Task batch multiplier should be 4 for '+$expectedLabels[$ti])}
    if((SN-Scalar-Int -value $taskData[$ti].partCount -default 0) -ne 1){throw ('Each thickness task should contain one part: '+$expectedLabels[$ti])}
  }
  $savedParts=@()
  foreach($rp in $parts){
    $found=SN-Find-WorkspacePartExact -app $verify -targetName $rp.part -sourcePath $rp.sourcePath -usedIndices @()
    if($null -eq $found){throw ('Mixed final workspace is missing '+$rp.part)}
    $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','Mat','MatName','MaterialType') -expectedText $rp.sigmaMaterial
    $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -expectedNumber $rp.thicknessMm
    $q=SN-Read-PartField -partObj $found.part -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $rp.qty
    if($null -eq $m -or $null -eq $t -or $null -eq $q){throw ('Mixed final CL values did not persist for '+$rp.part)}
    $savedParts += [pscustomobject]@{part=$rp.part;sourceType=$rp.sourceType;material=$m.value;thicknessMm=$t.value;qty=$q.value}
  }
  [pscustomobject]@{ok=$true;geometrySources='DXF+PRS';partCount=$partCount;taskCount=$taskCount;savedParts=$savedParts;taskData=$autoResult.taskData;workspaceBytes=(Get-Item -LiteralPath $wsPath).Length;timestamp=(Get-Date).ToString('o')}|ConvertTo-Json -Depth 10
}finally{
  foreach($o in @($verify,$verifyAuto)){if($o){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}catch{}}}
  Write-Host ('Mixed integration test workspace kept in TEMP for diagnosis: '+$wsPath)
}