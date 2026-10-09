$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
. (Join-Path $Root 'bridge\sigmanest-com.ps1')
$base='http://127.0.0.1:17832'
$prsRoot='S:\SNDataX1\PARTS'
$dxfRoot='Y:\'
$jobName='CLWS-HTTP-E2E-'+[Guid]::NewGuid().ToString('N').Substring(0,8)
$wsDirectory=Join-Path $env:TEMP $jobName
$stagingDirectory=Join-Path $prsRoot ('_CL_WS_BUILDER\'+$jobName)
$importJobId='';$autoJobId='';$wsPath='';$verify=$null
New-Item -ItemType Directory -Path $wsDirectory -Force|Out-Null

function Wait-Job([string]$id,[int]$timeoutSeconds=150){
  $until=(Get-Date).AddSeconds($timeoutSeconds)
  do{
    $s=Invoke-RestMethod -UseBasicParsing -TimeoutSec 5 -Uri ($base+'/api/build-status/'+$id)
    if([string]$s.state -in @('COMPLETE','FAILED')){return $s}
    Start-Sleep -Milliseconds 750
  }while((Get-Date) -lt $until)
  throw "Timed out waiting for bridge job $id."
}
try{
  $body=@{
    prsRoot=$prsRoot
    dxfRoot=$dxfRoot
    jobName=$jobName
    wsDirectory=$wsDirectory
    selectedSheetNames=@('SELFTEST')
    parts=@(@{
      part='PC-2A'
      qty=4
      material='Mild Steel'
      thickness='4'
      batchMultiplier=4
      sheet='SELFTEST'
      sourceSheets=@('SELFTEST')
      sourceRows=@()
      taskBatches=@()
    })
  }
  $accepted=Invoke-RestMethod -UseBasicParsing -Method Post -TimeoutSec 15 -Uri ($base+'/api/import-geometry') -ContentType 'application/json' -Body ($body|ConvertTo-Json -Depth 30)
  $importJobId=[string]$accepted.jobId
  if([string]::IsNullOrWhiteSpace($importJobId)){throw ('Import API did not return a jobId: '+($accepted|ConvertTo-Json -Depth 10))}
  $importStatus=Wait-Job -id $importJobId
  if($importStatus.state -ne 'COMPLETE'){throw ('Import API job failed: '+($importStatus|ConvertTo-Json -Depth 15))}
  if(-not $importStatus.result.ok){throw ('Import engine result was not OK: '+($importStatus.result|ConvertTo-Json -Depth 15))}
  $wsPath=[string]$importStatus.wsPath
  if([string]::IsNullOrWhiteSpace($wsPath)){ $wsPath=[string]$importStatus.result.wsPath }
  if(-not(Test-Path -LiteralPath $wsPath)){throw "Import API said complete but .ws file does not exist: $wsPath"}
  $importMethod=[string]$importStatus.result.parts[0].importMethod
  if($importMethod -ne 'LoadPart-DXF-GEOMETRY'){throw ('API import did not use the verified exact-path DXF loader: '+$importMethod)}
  $workerVersionImport=[string]$importStatus.workerVersion

  $autoBody=@{
    prsRoot=$prsRoot
    wsPath=$wsPath
    jobName=$jobName
    selectedSheetNames=@('SELFTEST')
    parts=@($importStatus.parts)
  }
  $autoAccepted=Invoke-RestMethod -UseBasicParsing -Method Post -TimeoutSec 15 -Uri ($base+'/api/autotask-label') -ContentType 'application/json' -Body ($autoBody|ConvertTo-Json -Depth 30)
  $autoJobId=[string]$autoAccepted.jobId
  if([string]::IsNullOrWhiteSpace($autoJobId)){throw ('AutoTask API did not return a jobId: '+($autoAccepted|ConvertTo-Json -Depth 10))}
  $autoStatus=Wait-Job -id $autoJobId
  if($autoStatus.state -ne 'COMPLETE'){throw ('AutoTask API job failed: '+($autoStatus|ConvertTo-Json -Depth 15))}
  if(-not $autoStatus.result.ok){throw ('AutoTask engine result was not OK: '+($autoStatus.result|ConvertTo-Json -Depth 15))}
  if(@($autoStatus.result.warnings|Where-Object{-not [string]::IsNullOrWhiteSpace([string]$_)}).Count -gt 0){
    throw ('AutoTask API returned warnings: '+(@($autoStatus.result.warnings) -join ' | '))
  }

  $verify=New-Object -ComObject SigmaNEST.SNApp
  [void]$verify.LoadWorkSpaceFile([string]$wsPath)
  $found=SN-Find-WorkspacePartExact -app $verify -targetName 'PC-2A' -sourcePath ([string]$importStatus.parts[0].sourcePath) -usedIndices @()
  if($null -eq $found){throw 'Final saved workspace does not contain PC-2A.'}
  $m=SN-Read-PartField -partObj $found.part -aliases @('Material','MaterialName','Mat','MatName','MaterialType') -expectedText 'MS'
  $t=SN-Read-PartField -partObj $found.part -aliases @('Thickness','SheetThickness','Thk','MaterialThickness','Thick') -expectedNumber 4
  $q=SN-Read-PartField -partObj $found.part -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt 4
  if($null -eq $m){throw 'Final API-created workspace material was not MS.'}
  if($null -eq $t){throw 'Final API-created workspace thickness was not 4mm.'}
  if($null -eq $q){throw 'Final API-created workspace Number To Nest was not 4.'}
  $taskCount=SN-Scalar-Int -value $verify.TasksList.Count -default 0
  if($taskCount -lt 1){throw 'Final API-created workspace has no saved task entries.'}
  $taskMatch=$null;$taskPart=$null;$taskIndex=-1
  for($ti=0;$ti -lt $taskCount;$ti++){
    $task=$null;try{$task=$verify.TasksList.Items($ti)}catch{continue}
    if($null -eq $task){continue}
    $pc=SN-Scalar-Int -value $task.PartsList.Count -default 0
    for($pi=0;$pi -lt $pc;$pi++){
      $tp=$null;try{$tp=$task.PartsList.Items($pi)}catch{continue}
      if($null -eq $tp){continue}
      $name='';try{$name=[string]$tp.Name}catch{}
      if($name.Equals('PC-2A',[StringComparison]::OrdinalIgnoreCase)){$taskMatch=$task;$taskPart=$tp;$taskIndex=$ti;break}
    }
    if($null -ne $taskMatch){break}
  }
  if($null -eq $taskMatch){throw 'No final task entry contains PC-2A.'}
  $expectedQty=16
  $tq=SN-Read-PartField -partObj $taskPart -aliases @('NumberToNest','NumberToLoad','QtyToNest','QuantityToNest','NestQuantity','QtyOrdered','Quantity','Qty','QtyRequired','QtyReq','QuantityOrdered','PartQuantity') -expectedInt $expectedQty
  if($null -eq $tq){throw 'Final task quantity is not 16 (CL qty 4 x batch 4).'}
  $label=('{0:00} | MS | 4mm' -f ($taskIndex+1))
  $savedLabel=''
  foreach($prop in @('Name','TaskName','Description')){try{$v=[string]$taskMatch.$prop;if($v.Trim().Equals($label,[StringComparison]::OrdinalIgnoreCase)){$savedLabel=$v;break}}catch{}}
  if(-not $savedLabel){throw ('Final task label was not saved as '+$label)}
  $batch=SN-Scalar-Int -value $taskMatch.BatchQuantity -default -1
  if($batch -ne 4){throw ('Final task BatchQuantity is not 4, got '+$batch)}

  [pscustomobject]@{
    ok=$true
    bridgeVersion=(Invoke-RestMethod -UseBasicParsing -TimeoutSec 5 -Uri ($base+'/api/health')).bridgeVersion
    importWorkerVersion=$workerVersionImport
    geometrySource='DXF'
    geometryPath=$importStatus.parts[0].sourcePath
    geometryMethod=$importMethod
    importedParts=$importStatus.importedCount
    tasksCreated=$autoStatus.result.tasksCreated
    savedTaskCount=$taskCount
    material=$m.value
    thicknessMm=$t.value
    partNumberToNest=$q.value
    taskLabel=$savedLabel
    taskBatchMultiplier=$batch
    taskQuantityWithBatch=$tq.value
    importJobId=$importJobId
    autoTaskJobId=$autoJobId
    workspace=$wsPath
    timestamp=(Get-Date).ToString('o')
  }|ConvertTo-Json -Depth 20
}finally{
  if($verify){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($verify)}catch{}}
  if($wsPath -and (Test-Path -LiteralPath $wsPath)){Remove-Item -LiteralPath $wsPath -Force -ErrorAction SilentlyContinue}
  if(Test-Path -LiteralPath $wsDirectory){Remove-Item -LiteralPath $wsDirectory -Recurse -Force -ErrorAction SilentlyContinue}
  if(Test-Path -LiteralPath $stagingDirectory){Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue}
  foreach($id in @($importJobId,$autoJobId)){
    if($id -match '^[0-9a-fA-F-]{36}$'){
      foreach($suffix in @('.json','.request.json')){
        $file=Join-Path (Join-Path $Root 'bridge\build-status') ($id+$suffix)
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
      }
    }
  }
}