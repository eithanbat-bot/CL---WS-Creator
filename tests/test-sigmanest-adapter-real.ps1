param(
  [Parameter(Mandatory=$true)][string]$WorkspacePath,
  [Parameter(Mandatory=$true)][string]$PartName,
  [int]$TestQty=4,
  [string]$TestMaterial='MS',
  [double]$TestThicknessMm=4,
  [int]$BatchMultiplier=4
)
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
. (Join-Path $Root 'bridge\sigmanest-com.ps1')
if(-not(Test-Path -LiteralPath $WorkspacePath)){throw "Workspace not found: $WorkspacePath"}
if($TestQty -lt 1){throw 'TestQty must be at least 1.'}
if($BatchMultiplier -lt 1){throw 'BatchMultiplier must be at least 1.'}

$outDir=Join-Path $env:TEMP 'CL-WS-Creator-AdapterSelfTest'
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$copy=Join-Path $outDir ('AUTOTASK-SELFTEST-'+[Guid]::NewGuid().ToString('N')+'.ws')
Copy-Item -LiteralPath $WorkspacePath -Destination $copy -Force
$verifyApp=$null
try{
  $requestPart=[pscustomobject]@{
    part=$PartName
    qty=$TestQty
    sigmaMaterial=$TestMaterial
    thicknessMm=$TestThicknessMm
    # AutoTask works on already imported workspace geometry; this path is used
    # only to corroborate the part identity in the isolated test copy.
    sourcePath=('C:\CL-WS-Creator-SELFTEST\'+$PartName+'.DXF')
    sourceType='DXF'
    batchMultiplier=$BatchMultiplier
    taskBatches=@()
    sourceSheets=@()
    sourceRows=@()
  }
  $request=[pscustomobject]@{wsPath=$copy;jobName='CL_WS_CREATOR_SELFTEST';parts=@($requestPart)}

  # Exactly one result object must leave the adapter. A COM collection or
  # other pipeline value here caused the previous System.Object[] bug.
  $rawResults=@(Invoke-SigmaNestAutoTask -Request $request)
  if($rawResults.Count -ne 1){
    $types=@($rawResults|ForEach-Object{try{$_.GetType().FullName}catch{'<unknown>'}})
    throw ('AutoTask adapter emitted '+$rawResults.Count+' pipeline values instead of exactly one result. Types: '+($types -join ', '))
  }
  $result=$rawResults[0]
  if($null -eq $result -or $result.PSObject.Properties.Name -notcontains 'ok' -or $result.PSObject.Properties.Name -notcontains 'phase'){
    throw 'AutoTask adapter output is not the expected structured result object.'
  }
  if(-not $result.ok){throw ('Adapter returned not-ok at '+$result.phase+': '+$result.message+'; warnings='+( @($result.warnings) -join ' | '))}
  $realWarnings=@($result.warnings | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
  if($realWarnings.Count -gt 0){throw ('Adapter returned warnings: '+($realWarnings -join ' | '))}
  if(@($result.taskData).Count -lt 1){throw 'Adapter produced no task-data confirmation.'}
  $taskDetail=@($result.taskData | Where-Object {$_.labelApplied -and $_.batchApplied})
  if($taskDetail.Count -lt 1){throw 'Adapter did not confirm a task label and batch multiplier.'}

  # Reload only the temporary saved workspace to validate persisted values.
  $verifyApp=New-Object -ComObject SigmaNEST.SNApp
  [void]$verifyApp.LoadWorkSpaceFile([string]$copy)
  $found=SN-Find-WorkspacePartExact -app $verifyApp -targetName $PartName -sourcePath $requestPart.sourcePath -usedIndices @()
  if($null -eq $found){throw "Saved workspace verification could not find part: $PartName"}
  $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','Mat','MatName','MaterialType') -expectedText $TestMaterial
  if($null -eq $m){throw "Saved part material is not $TestMaterial"}
  $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -expectedNumber $TestThicknessMm
  if($null -eq $t){throw "Saved part thickness is not $TestThicknessMm mm"}
  $q=SN-Read-PartField -partObj $found.part -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $TestQty
  if($null -eq $q){throw "Saved Part Parameters quantity is not $TestQty"}

  $taskCount=SN-Scalar-Int -value $verifyApp.TasksList.Count -default 0
  if($taskCount -lt 1){throw 'Saved workspace has no TasksList entries.'}
  $taskMatch=$null;$taskPartMatch=$null;$taskIndex=-1
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$verifyApp.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $pc=SN-Scalar-Int -value $task.PartsList.Count -default 0
    for($pi=0;$pi -lt $pc;$pi++){
      $tp=$null;try{$tp=$task.PartsList.Items($pi)}catch{continue}
      if($null -eq $tp){continue}
      $name='';try{$name=[string]$tp.Name}catch{}
      if($name.Equals($PartName,[StringComparison]::OrdinalIgnoreCase)){
        $taskMatch=$task;$taskPartMatch=$tp;$taskIndex=$ti;break
      }
    }
    if($null -ne $taskMatch){break}
  }
  if($null -eq $taskMatch){throw "Saved TasksList does not contain test part '$PartName'."}

  $expectedTaskQty=$TestQty*$BatchMultiplier
  $taskQty=SN-Read-PartField -partObj $taskPartMatch -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $expectedTaskQty
  if($null -eq $taskQty){throw "Saved task quantity is not $expectedTaskQty (CL qty $TestQty x batch $BatchMultiplier)."}

  $expectedLabel=('{0:00} | {1} | {2:0.###}mm' -f ($taskIndex+1),$TestMaterial,$TestThicknessMm)
  $label=''
  foreach($prop in @('Name','TaskName','Description')){
    try{
      $value=[string]$taskMatch.$prop
      if($value.Trim().Equals($expectedLabel.Trim(),[StringComparison]::OrdinalIgnoreCase)){$label=$value;break}
    }catch{}
  }
  if([string]::IsNullOrWhiteSpace($label)){throw "Saved task label verification failed; expected '$expectedLabel'."}

  $batch=-1;$batchProperty=''
  foreach($prop in @('BatchMultiplier','BatchQty','BatchQuantity','Batch')){
    try{
      $v=SN-Scalar-Int -value $taskMatch.$prop -default -1
      if($v -eq $BatchMultiplier){$batch=$v;$batchProperty=$prop;break}
    }catch{}
  }
  if($batch -ne $BatchMultiplier){
    foreach($prop in @('BatchQty','BatchMultiplier','BatchQuantity','Batch')){
      try{
        $v=SN-Scalar-Int -value $taskPartMatch.$prop -default -1
        if($v -eq $BatchMultiplier){$batch=$v;$batchProperty='PART.'+$prop;break}
      }catch{}
    }
  }
  if($batch -ne $BatchMultiplier){throw "Saved task batch verification failed; expected multiplier $BatchMultiplier."}

  [pscustomobject]@{
    ok=$true
    sourceWorkspace=$WorkspacePath
    part=$PartName
    persistedMaterial=$m.value
    persistedThickness=$t.value
    partNumberToNest=$q.value
    tasks=$taskCount
    taskLabel=$label
    taskBatchMultiplier=$batch
    batchProperty=$batchProperty
    taskQuantity=$taskQty.value
    expectedTaskQuantity=$expectedTaskQty
    adapterResult=$result.message
    timestamp=(Get-Date).ToString('o')
  } | ConvertTo-Json -Depth 20
}finally{
  if($verifyApp){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($verifyApp)}catch{}}
  if(Test-Path -LiteralPath $copy){Remove-Item -LiteralPath $copy -Force -ErrorAction SilentlyContinue}
}
