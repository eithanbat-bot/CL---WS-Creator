param(
  [Parameter(Mandatory=$true)][string]$RequestFile
)

$ErrorActionPreference='Stop'
$WorkerVersion='2.12.2'
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
    $data=Invoke-SigmaNestAutoTask -Request $request.autoTaskRequest
  }else{
    $phase='PREPARE'
    Write-BuildStatus $statusFile ([ordered]@{jobId=[string]$request.jobId;state='RUNNING';phase=$phase;message='Matching CL parts and preparing the SigmaNEST geometry operation.';workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$null;elapsedSeconds=((Get-Date).ToUniversalTime()-$started).TotalSeconds;result=$null;parts=@($request.reportParts);jobName=[string]$request.jobName;selectedSheets=@($request.selectedSheets)})
    $prepared=Prepare-SigmaNestBuild -b $request.buildRequest
    if($null -eq $prepared){throw 'Background build preparation returned no result.'}
    if($prepared.PSObject.Properties.Name -contains 'Status' -and [int]$prepared.Status -ne 200){throw [string]$prepared.Data.message}
    $request.reportParts=@($prepared.parts)
    $request.outputDir=[string]$prepared.outputDir
    $request.jobName=[string]$prepared.jobName
    $request.selectedSheets=@($prepared.selectedSheets)

    if($prepared.noGeometry){
      $finished=(Get-Date).ToUniversalTime()
      Write-BuildStatus $statusFile ([ordered]@{jobId=[string]$request.jobId;state='COMPLETE';phase='PREPARE';message=[string]$prepared.message;workerVersion=$WorkerVersion;pid=$PID;started=$started.ToString('o');finished=$finished.ToString('o');elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1);result=$prepared;parts=@($prepared.parts);outputDir=[string]$prepared.outputDir;jobName=[string]$prepared.jobName;selectedSheets=@($prepared.selectedSheets);reviewCount=[int]$prepared.reviewCount;reviewBreakdown=$prepared.reviewBreakdown;importedCount=0;missingCount=[int]$prepared.missingCount;wsPath=''}) 
      Remove-Item -LiteralPath $RequestFile -Force -ErrorAction SilentlyContinue
      exit 0
    }

    $engineRequest=$prepared.engineRequest
    if($null -eq $engineRequest){throw 'Background preparation returned no SigmaNEST engine request.'}
    $phase=if($mode -eq 'IMPORT_ONLY'){'IMPORT_GEOMETRY'}else{'COM_BUILD'}
    $data=if($mode -eq 'IMPORT_ONLY'){Invoke-SigmaNestImportGeometry -Request $engineRequest}else{Invoke-SigmaNestBuild -Request $engineRequest}
  }

  if($null -eq $data){throw 'Background SigmaNEST engine returned no result.'}
  $finished=(Get-Date).ToUniversalTime()
  $state=if([bool]$data.ok){'COMPLETE'}else{'FAILED'}
  $message=if([bool]$data.ok){[string]$data.message}else{[string]$data.error}
  $finalParts=if($mode -eq 'IMPORT_ONLY'){@($request.reportParts)}elseif($data.parts){@($data.parts)}else{@($request.reportParts)}
  $finalOutput=if($data.outputDir){[string]$data.outputDir}else{[string]$request.outputDir}
  $finalWs=[string]$data.wsPath
  $reviewCount=if($prepared){[int]$prepared.reviewCount}else{0}
  $reviewBreakdown=if($prepared){$prepared.reviewBreakdown}else{@{}}
  $imported=if($mode -eq 'IMPORT_ONLY' -and $data.partCount -ne $null){[int]$data.partCount}elseif($data.partCount -ne $null){[int]$data.partCount}else{0}
  $missing=if($prepared){[int]$prepared.missingCount}else{0}
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
