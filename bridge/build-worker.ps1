param(
  [Parameter(Mandatory=$true)][string]$RequestFile
)

$ErrorActionPreference='Stop'
$WorkerVersion='2.24.0'
$Root=Split-Path -Parent $MyInvocation.MyCommand.Path
$ComLibrary=Join-Path $Root 'sigmanest-com.ps1'

function Write-BuildStatus($statusFile,$obj){
  $dir=Split-Path -Parent $statusFile
  New-Item -ItemType Directory -Path $dir -Force|Out-Null
  $tmp=$statusFile+'.tmp-'+[Guid]::NewGuid().ToString('N')
  ($obj|ConvertTo-Json -Depth 30)|Set-Content -LiteralPath $tmp -Encoding UTF8
  Move-Item -LiteralPath $tmp -Destination $statusFile -Force
}

$request=$null
$statusFile=$null
$prepared=$null
$started=(Get-Date).ToUniversalTime()
try{
  if(-not(Test-Path -LiteralPath $RequestFile)){throw 'Build request file not found: '+$RequestFile}
  $request=Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8|ConvertFrom-Json
  $statusFile=[string]$request.statusFile
  if([string]::IsNullOrWhiteSpace($statusFile)){throw 'Build request does not contain statusFile.'}
  $mode=[string]$request.mode
  if([string]::IsNullOrWhiteSpace($mode)){$mode='FULL'}

  Write-BuildStatus $statusFile ([ordered]@{jobId=[string]$request.jobId;state='RUNNING';phase='START';message=('Background operation started: '+$mode);workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$null;elapsedSeconds=0;result=$null;parts=@($request.reportParts);outputDir=[string]$request.outputDir;jobName=[string]$request.jobName;selectedSheets=@($request.selectedSheets)})

  if(-not(Test-Path -LiteralPath $ComLibrary)){throw 'SigmaNEST COM library is missing: '+$ComLibrary}
  . $ComLibrary
  $serverLibrary=Join-Path $Root 'server.ps1'
  if(-not(Test-Path -LiteralPath $serverLibrary)){throw 'Bridge server library is missing: '+$serverLibrary}
  . $serverLibrary -LibraryOnly

  if($mode -eq 'AUTOTASK_ONLY'){
    $phase='AUTOTASK'
    Write-BuildStatus $statusFile ([ordered]@{jobId=[string]$request.jobId;state='RUNNING';phase=$phase;message='Applying CL part data, running AutoTask and applying batch/order labels.';workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$null;elapsedSeconds=((Get-Date).ToUniversalTime()-$started).TotalSeconds;result=$null;parts=@($request.reportParts);outputDir=[string]$request.outputDir;jobName=[string]$request.jobName;selectedSheets=@($request.selectedSheets)})
    # SigmaNEST COM methods can emit non-result pipeline objects. Keep the
    # AutoTask worker contract identical to the geometry-import contract:
    # only the final structured result may reach the status/report layer.
    $rawAutoTaskData=@(Invoke-SigmaNestAutoTask -Request $request.autoTaskRequest)
    $autoTaskResults=@($rawAutoTaskData | Where-Object {
      $null -ne $_ -and
      $_.PSObject.Properties.Name -contains 'ok' -and
      $_.PSObject.Properties.Name -contains 'phase'
    })
    if($autoTaskResults.Count -eq 0){
      $types=@($rawAutoTaskData|ForEach-Object{try{$_.GetType().FullName}catch{'<unknown>'}})
      throw ('AutoTask engine returned no structured result. Raw pipeline types: '+($types -join ', '))
    }
    $data=$autoTaskResults[-1]
  }else{
    $phase='PREPARE'
    Write-BuildStatus $statusFile ([ordered]@{jobId=[string]$request.jobId;state='RUNNING';phase=$phase;message='Matching CL parts and preparing the SigmaNEST geometry operation.';workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$null;elapsedSeconds=((Get-Date).ToUniversalTime()-$started).TotalSeconds;result=$null;parts=@($request.reportParts);jobName=[string]$request.jobName;selectedSheets=@($request.selectedSheets)})
    $prepared=Prepare-SigmaNestBuild -b $request.buildRequest
    if($null -eq $prepared){throw 'Background build preparation returned no result.'}
    if($prepared.PSObject.Properties.Name -contains 'Status' -and (SN-Scalar-Int -value $prepared.Status -default -1) -ne 200){throw [string]$prepared.Data.message}
    $request.reportParts=@($prepared.parts)
    $request.outputDir=[string]$prepared.outputDir
    $request.jobName=[string]$prepared.jobName
    $request.selectedSheets=@($prepared.selectedSheets)

    if($prepared.noGeometry){
      $finished=(Get-Date).ToUniversalTime()
      Write-BuildStatus $statusFile ([ordered]@{jobId=[string]$request.jobId;state='COMPLETE';phase='PREPARE';message=[string]$prepared.message;workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$finished.ToString('o');elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1);result=$prepared;parts=@($prepared.parts);outputDir=[string]$prepared.outputDir;jobName=[string]$prepared.jobName;selectedSheets=@($prepared.selectedSheets);reviewCount=(SN-Scalar-Int -value $prepared.reviewCount -default 0);reviewBreakdown=$prepared.reviewBreakdown;importedCount=0;missingCount=(SN-Scalar-Int -value $prepared.missingCount -default 0);wsPath=''}) 
      Remove-Item -LiteralPath $RequestFile -Force -ErrorAction SilentlyContinue
      exit 0
    }

    $engineRequest=$prepared.engineRequest
    if($null -eq $engineRequest){throw 'Background preparation returned no SigmaNEST engine request.'}
    $phase=if($mode -eq 'IMPORT_ONLY'){'IMPORT_GEOMETRY'}else{'COM_BUILD'}

    # SigmaNEST COM calls can occasionally emit an extra pipeline value even
    # when the operation itself succeeds. Capture the complete output and take
    # only the structured engine result object. This prevents an accidental
    # System.Object[] from reaching the Int32/count handling below.
    $rawEngineData=@(if($mode -eq 'IMPORT_ONLY'){
      Invoke-SigmaNestImportGeometry -Request $engineRequest
    }else{
      Invoke-SigmaNestBuild -Request $engineRequest
    })
    $engineResults=@($rawEngineData | Where-Object {
      $null -ne $_ -and
      $_.PSObject.Properties.Name -contains 'ok' -and
      $_.PSObject.Properties.Name -contains 'phase'
    })
    if($engineResults.Count -eq 0){
      $types=@($rawEngineData|ForEach-Object{try{$_.GetType().FullName}catch{'<unknown>'}})
      throw ('Background SigmaNEST engine returned no structured result. Raw pipeline types: '+($types -join ', '))
    }
    $data=$engineResults[-1]
  }

  if($null -eq $data){throw 'Background SigmaNEST engine returned no result.'}
  $finished=(Get-Date).ToUniversalTime()
  $state=if([bool]$data.ok){'COMPLETE'}else{'FAILED'}
  $message=if([bool]$data.ok){[string]$data.message}else{
    if($data.phase -and $data.error){"$($data.phase): $($data.error)"}else{[string]$data.error}
  }
  $finalParts=if($mode -eq 'IMPORT_ONLY'){@($request.reportParts)}elseif($data.parts){@($data.parts)}else{@($request.reportParts)}
  $finalOutput=if($data.outputDir){[string]$data.outputDir}else{[string]$request.outputDir}
  $finalWs=[string]$data.wsPath
  $reviewCount=if($prepared){SN-Scalar-Int -value $prepared.reviewCount -default 0}else{0}
  $reviewBreakdown=if($prepared){$prepared.reviewBreakdown}else{@{}}
  $imported=if($data.partCount -ne $null){SN-Scalar-Int -value $data.partCount -default 0}else{0}
  $missing=if($prepared){SN-Scalar-Int -value $prepared.missingCount -default 0}else{0}
  $extra=[ordered]@{
    jobId=[string]$request.jobId;state=$state;phase=[string]$data.phase;message=$message;workerVersion=$WorkerVersion;pid=$PID
    started=$started.ToString('o');finished=$finished.ToString('o');elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1)
    result=$data;parts=$finalParts;outputDir=$finalOutput;jobName=[string]$request.jobName;selectedSheets=@($request.selectedSheets)
    reviewCount=$reviewCount;reviewBreakdown=$reviewBreakdown;importedCount=$imported;missingCount=$missing;wsPath=$finalWs
  }
  Write-BuildStatus $statusFile $extra
  Remove-Item -LiteralPath $RequestFile -Force -ErrorAction SilentlyContinue
}catch{
  $finished=(Get-Date).ToUniversalTime()
  $err=$_.Exception.Message
  try{
    if(-not $statusFile -and $request){$statusFile=[string]$request.statusFile}
    if($statusFile){
      Write-BuildStatus $statusFile ([ordered]@{jobId=if($request){[string]$request.jobId}else{''};state='FAILED';phase='WORKER';message=$err;error=$err;workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$finished.ToString('o');elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1);result=$null;parts=if($request){@($request.reportParts)}else{@()};outputDir=if($request){[string]$request.outputDir}else{''};jobName=if($request){[string]$request.jobName}else{''}})
    }
  }catch{}
  exit 1
}
