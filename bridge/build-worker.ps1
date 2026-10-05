param(
  [Parameter(Mandatory=$true)][string]$RequestFile
)

$ErrorActionPreference='Stop'
$WorkerVersion='2.11.2'
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
$started=(Get-Date).ToUniversalTime()
try{
  if(-not(Test-Path -LiteralPath $RequestFile)){throw 'Build request file not found: '+$RequestFile}
  $request=Get-Content -LiteralPath $RequestFile -Raw -Encoding UTF8|ConvertFrom-Json
  $statusFile=[string]$request.statusFile
  if([string]::IsNullOrWhiteSpace($statusFile)){throw 'Build request does not contain statusFile.'}

  Write-BuildStatus $statusFile ([ordered]@{
    jobId=[string]$request.jobId
    state='RUNNING'
    phase='START'
    message='SigmaNEST background worker started.'
    workerVersion=$WorkerVersion
    pid=$PID
    started=$started.ToString('o')
    finished=$null
    elapsedSeconds=0
    result=$null
    parts=@($request.reportParts)
  })

  $serverLibrary=Join-Path $Root 'server.ps1'
  if(-not(Test-Path -LiteralPath $serverLibrary)){throw 'Bridge server library is missing: '+$serverLibrary}
  . $serverLibrary -LibraryOnly

  Write-BuildStatus $statusFile ([ordered]@{
    jobId=[string]$request.jobId
    state='RUNNING'
    phase='PREPARE'
    message='Matching CL parts, validating geometry and staging the SigmaNEST job.'
    workerVersion=$WorkerVersion
    pid=$PID
    started=$started.ToString('o')
    finished=$null
    elapsedSeconds=((Get-Date).ToUniversalTime()-$started).TotalSeconds
    result=$null
    parts=@($request.reportParts)
    jobName=[string]$request.jobName
    selectedSheets=@($request.selectedSheets)
  })

  $prepared=Prepare-SigmaNestBuild -b $request.buildRequest
  if($null -eq $prepared){throw 'Background build preparation returned no result.'}
  $request.reportParts=@($prepared.parts)
  $request.outputDir=[string]$prepared.outputDir
  $request.jobName=[string]$prepared.jobName
  $request.selectedSheets=@($prepared.selectedSheets)

  if($prepared.noGeometry){
    $finished=(Get-Date).ToUniversalTime()
    Write-BuildStatus $statusFile ([ordered]@{
      jobId=[string]$request.jobId
      state='COMPLETE'
      phase='PREPARE'
      message=[string]$prepared.message
      workerVersion=$WorkerVersion
      pid=$PID
      started=$started.ToString('o')
      finished=$finished.ToString('o')
      elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1)
      result=$prepared
      parts=@($prepared.parts)
      outputDir=[string]$prepared.outputDir
      jobName=[string]$prepared.jobName
      selectedSheets=@($prepared.selectedSheets)
      reviewCount=[int]$prepared.reviewCount
      reviewBreakdown=$prepared.reviewBreakdown
      importedCount=[int]$prepared.importedCount
      missingCount=[int]$prepared.missingCount
      wsPath=''
    })
    Remove-Item -LiteralPath $RequestFile -Force -ErrorAction SilentlyContinue
    exit 0
  }

  $request.engineRequest=$prepared.engineRequest
  $phase='COM_BUILD'
  Write-BuildStatus $statusFile ([ordered]@{
    jobId=[string]$request.jobId
    state='RUNNING'
    phase=$phase
    message='SigmaNEST is importing geometry and constructing the workspace.'
    workerVersion=$WorkerVersion
    pid=$PID
    started=$started.ToString('o')
    finished=$null
    elapsedSeconds=((Get-Date).ToUniversalTime()-$started).TotalSeconds
    result=$null
    parts=@($request.reportParts)
  })

  $data=Invoke-SigmaNestBuild -Request $request.engineRequest
  $finished=(Get-Date).ToUniversalTime()
  if($null -eq $data){throw 'SigmaNEST background COM engine returned no result.'}

  $state=if([bool]$data.ok){'COMPLETE'}else{'FAILED'}
  $message=if([bool]$data.ok){[string]$data.message}else{[string]$data.error}
  Write-BuildStatus $statusFile ([ordered]@{
    jobId=[string]$request.jobId
    state=$state
    phase=[string]$data.phase
    message=$message
    workerVersion=$WorkerVersion
    pid=$PID
    started=$started.ToString('o')
    finished=$finished.ToString('o')
    elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1)
    result=$data
    parts=@($request.reportParts)
    outputDir=[string]$request.outputDir
    jobName=[string]$request.jobName
    selectedSheets=@($request.selectedSheets)
    reviewCount=$(if($null -ne $prepared.reviewCount){[int]$prepared.reviewCount}else{0})
    reviewBreakdown=$(if($null -ne $prepared.reviewBreakdown){$prepared.reviewBreakdown}else{@{}})
    importedCount=$(if($null -ne $data.importedCount){[int]$data.importedCount}else{[int]$data.partCount})
    missingCount=$(if($null -ne $data.missingCount){[int]$data.missingCount}else{[int]$prepared.missingCount})
    wsPath=[string]$data.wsPath
  })
}catch{
  $finished=(Get-Date).ToUniversalTime()
  $err=$_.Exception.Message
  try{
    if(-not $statusFile -and $request){$statusFile=[string]$request.statusFile}
    if($statusFile){
      Write-BuildStatus $statusFile ([ordered]@{
        jobId=if($request){[string]$request.jobId}else{''}
        state='FAILED'
        phase='WORKER'
        message=$err
        error=$err
        workerVersion=$WorkerVersion
        pid=$PID
        started=$started.ToString('o')
        finished=$finished.ToString('o')
        elapsedSeconds=[math]::Round((($finished-$started).TotalSeconds),1)
        result=$null
        parts=if($request){@($request.reportParts)}else{@()}
        outputDir=if($request){[string]$request.outputDir}else{''}
        jobName=if($request){[string]$request.jobName}else{''}
      })
    }
  }catch{}
  exit 1
}
