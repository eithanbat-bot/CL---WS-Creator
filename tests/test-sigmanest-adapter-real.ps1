param(
  [string]$DxfPath='Y:\AutoCAD\Hino old\Hino 300\Rear Bodies\SBV\2023 (HSW) SBV Hino 300 816 Body\Dxf Files\PC-2A.DXF',
  [int]$TestQty=4,
  [string]$TestMaterial='MS',
  [double]$TestThicknessMm=4.0,
  [int]$BatchMultiplier=4
)
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
. (Join-Path $Root 'bridge\sigmanest-com.ps1')
if(-not(Test-Path -LiteralPath $DxfPath)){throw "DXF test geometry not found: $DxfPath"}
if([IO.Path]::GetExtension($DxfPath) -ine '.dxf'){throw "The real adapter self-test requires a .DXF source: $DxfPath"}
if($TestQty -lt 1){throw 'TestQty must be at least 1.'}
if($BatchMultiplier -lt 1){throw 'BatchMultiplier must be at least 1.'}
$partName=[IO.Path]::GetFileNameWithoutExtension($DxfPath)
$outDir=Join-Path $env:TEMP 'CL-WS-Creator-FullE2ETest'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$jobName='CLWS-E2E-'+[Guid]::NewGuid().ToString('N').Substring(0,10)
$importResult=$null;$autoResult=$null;$verify=$null;$verifyAutomation=$null;$wsPath=''
try{
  $rp=[pscustomobject]@{
    part=$partName;qty=$TestQty;sigmaMaterial=$TestMaterial;thicknessMm=$TestThicknessMm
    sourcePath=$DxfPath;sourceType='DXF';matchType='SELFTEST'
    batchMultiplier=$BatchMultiplier;taskBatches=@();sourceSheets=@();sourceRows=@()
  }
  $importRequest=[pscustomobject]@{
    jobName=$jobName;wsDirectory=$outDir;libraryRoot='S:\SNDataX1\PARTS'
    clLinkFile='';parts=@($rp)
  }
  $rawImport=@(Invoke-SigmaNestImportGeometry -Request $importRequest)
  if($rawImport.Count -ne 1){throw ('Import returned '+$rawImport.Count+' pipeline values; expected exactly one.')}
  $importResult=$rawImport[0]
  if(-not $importResult.ok){throw ('Import failed at '+$importResult.phase+': '+$importResult.error)}
  if([string]$importResult.parts[0].sourceType -ne 'DXF'){throw 'Import result did not report DXF as source type.'}
  if([string]$importResult.parts[0].sourcePath -ne $DxfPath){throw 'Import result path differs from the selected DXF.'}
  if([string]$importResult.parts[0].importMethod -ne 'AddPartImport-DXF-GEOMETRY'){throw ('Unexpected geometry method: '+[string]$importResult.parts[0].importMethod)}
  $wsPath=[string]$importResult.wsPath
  if(-not(Test-Path -LiteralPath $wsPath)){throw 'DXF import did not create the new WS file.'}

  $autoRequest=[pscustomobject]@{wsPath=$wsPath;jobName=$jobName;parts=@($rp)}
  $rawAuto=@(Invoke-SigmaNestAutoTask -Request $autoRequest)
  if($rawAuto.Count -ne 1){
    $types=@($rawAuto|ForEach-Object{try{$_.GetType().FullName}catch{'<unknown>'}})
    throw ('AutoTask emitted '+$rawAuto.Count+' pipeline values; expected exactly one. Types: '+($types -join ', '))
  }
  $autoResult=$rawAuto[0]
  if($null -eq $autoResult -or $autoResult.PSObject.Properties.Name -notcontains 'ok'){throw 'AutoTask returned no structured result.'}
  if(-not $autoResult.ok){throw ('AutoTask failed at '+$autoResult.phase+': '+$autoResult.message+'; warnings=' + (@($autoResult.warnings) -join ' | '))}
  $actualWarnings=@($autoResult.warnings|Where-Object{-not [string]::IsNullOrWhiteSpace([string]$_)})
  if($actualWarnings.Count -gt 0){throw ('AutoTask warnings: '+($actualWarnings -join ' | '))}
  if((SN-Scalar-Int -value $autoResult.tasksCreated -default 0) -lt 1){throw 'AutoTask did not report any created tasks.'}
  if(@($autoResult.taskData).Count -lt 1){throw 'AutoTask did not return task labels/batch confirmations.'}
  if(@($autoResult.taskData|Where-Object{$_.labelApplied -and $_.batchApplied -and [int]$_.batchMultiplier -eq $BatchMultiplier}).Count -lt 1){throw "AutoTask did not confirm label and batch multiplier x$BatchMultiplier."}

  $verifyAutomation=New-Object -ComObject SigmaNEST.SNAutomation
  [void]$verifyAutomation.FileNew()
  Start-Sleep -Milliseconds 250
  $verify=New-Object -ComObject SigmaNEST.SNApp
  [void]$verify.LoadWorkSpaceFile([string]$wsPath)
  Start-Sleep -Milliseconds 250
  $found=SN-Find-WorkspacePartExact -app $verify -targetName $partName -sourcePath $DxfPath -usedIndices @()
  if($null -eq $found){throw "Saved workspace does not contain the selected DXF part $partName."}
  $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','Mat','MatName','MaterialType') -expectedText $TestMaterial
  $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -expectedNumber $TestThicknessMm
  $q=SN-Read-PartField -partObj $found.part -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $TestQty
  if($null -eq $m){throw "Reloaded part material was not $TestMaterial."}
  if($null -eq $t){throw "Reloaded part thickness was not $TestThicknessMm mm."}
  if($null -eq $q){throw "Reloaded Number To Nest was not $TestQty."}
  $taskCount=SN-Scalar-Int -value $verify.TasksList.Count -default 0
  if($taskCount -lt 1){throw 'Saved workspace did not persist any TasksList entries.'}
  $taskMatch=$null;$taskPartMatch=$null;$taskIndex=-1
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$verify.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $pc=SN-Scalar-Int -value $task.PartsList.Count -default 0
    for($pi=0;$pi -lt $pc;$pi++){
      $tp=$null;try{$tp=$task.PartsList.Items($pi)}catch{continue}
      if($null -eq $tp){continue}
      $name='';try{$name=[string]$tp.Name}catch{}
      if($name.Equals($partName,[StringComparison]::OrdinalIgnoreCase)){
        $taskMatch=$task;$taskPartMatch=$tp;$taskIndex=$ti;break
      }
    }
    if($null -ne $taskMatch){break}
  }
  if($null -eq $taskMatch){throw "No saved task contains $partName."}
  $expectedTaskQty=$TestQty*$BatchMultiplier
  $tq=SN-Read-PartField -partObj $taskPartMatch -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $expectedTaskQty
  if($null -eq $tq){throw "Saved task quantity was not $expectedTaskQty (CL qty $TestQty x batch $BatchMultiplier)."}
  $expectedLabel=('{0:00} | {1} | {2:0.###}mm' -f ($taskIndex+1),$TestMaterial,$TestThicknessMm)
  $label=''
  foreach($prop in @('Name','TaskName','Description')){
    try{$v=[string]$taskMatch.$prop;if($v.Trim().Equals($expectedLabel,[StringComparison]::OrdinalIgnoreCase)){$label=$v;break}}catch{}
  }
  if([string]::IsNullOrWhiteSpace($label)){throw "Saved task label did not match $expectedLabel."}
  $batch=-1;$batchProp=''
  foreach($prop in @('BatchMultiplier','BatchQty','BatchQuantity','Batch')){
    try{$v=SN-Scalar-Int -value $taskMatch.$prop -default -1;if($v -eq $BatchMultiplier){$batch=$v;$batchProp=$prop;break}}catch{}
  }
  if($batch -ne $BatchMultiplier){
    foreach($prop in @('BatchQty','BatchMultiplier','BatchQuantity','Batch')){
      try{$v=SN-Scalar-Int -value $taskPartMatch.$prop -default -1;if($v -eq $BatchMultiplier){$batch=$v;$batchProp='PART.'+$prop;break}}catch{}
    }
  }
  if($batch -ne $BatchMultiplier){throw "Saved task batch multiplier was not x$BatchMultiplier."}

  [pscustomobject]@{
    ok=$true
    geometrySource='DXF';geometryPath=$importResult.parts[0].sourcePath;geometryMethod=$importResult.parts[0].importMethod
    workspacePath=$wsPath;part=$partName;importedPartCount=$importResult.partCount
    tasksCreated=$autoResult.tasksCreated;savedTaskCount=$taskCount
    partMaterial=$m.value;partThicknessMm=$t.value;partNumberToNest=$q.value
    taskLabel=$label;taskBatchMultiplier=$batch;batchProperty=$batchProp
    taskQtyWithBatch=$tq.value;expectedTaskQty=$expectedTaskQty
    importPipelineValues=$rawImport.Count;autoTaskPipelineValues=$rawAuto.Count
    timestamp=(Get-Date).ToString('o')
  }|ConvertTo-Json -Depth 20
}finally{
  if($verify){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($verify)}catch{}}
  if($verifyAutomation){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($verifyAutomation)}catch{}}
  if($wsPath -and (Test-Path -LiteralPath $wsPath)){Remove-Item -LiteralPath $wsPath -Force -ErrorAction SilentlyContinue}
}