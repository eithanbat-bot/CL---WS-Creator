# Static release-gate checks for the SigmaNEST adapter.
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
$server=Get-Content -Raw (Join-Path $Root 'bridge\server.ps1')
$worker=Get-Content -Raw (Join-Path $Root 'bridge\build-worker.ps1')
$com=Get-Content -Raw (Join-Path $Root 'bridge\sigmanest-com.ps1')
$taskpane=Get-Content -Raw (Join-Path $Root 'taskpane.js')

function Assert([bool]$condition,[string]$message){if(-not $condition){throw $message}}

$serverVersion=[regex]::Match($server,'(?m)^\s*\$BRIDGE_VERSION\s*=\s*''([^'']+)''').Groups[1].Value
$workerVersion=[regex]::Match($worker,'(?m)^\s*\$WorkerVersion\s*=\s*''([^'']+)''').Groups[1].Value
Assert ($serverVersion -eq $workerVersion) "Bridge/worker version mismatch: $serverVersion vs $workerVersion"
Assert ($serverVersion -eq '2.19.0') "Unexpected bridge version: $serverVersion"

Assert ($com.Contains("function Invoke-SigmaNestImportGeometry")) 'Import entry point missing.'
Assert ($com.Contains("function Invoke-SigmaNestAutoTask")) 'AutoTask entry point missing.'
Assert ($com.Contains("CreateTasksListForNewPartsInWS")) 'Task creation method missing.'
Assert ($com -match '\$phase\s*=\s*''CREATE_TASKS_FOR_IMPORTED_PARTS''') 'AutoTask phase does not explicitly create tasks before AutoTask.'
Assert ($com -match '\$phase\s*=\s*''AUTO_TASK''') 'AutoTask phase missing.'
Assert ($com.Contains('try{$app.RefreshTreeView()}catch{}')) 'SigmaNEST tree refresh missing.'
Assert ($com.Contains("NumberToNest','NumberToLoad'")) 'Part quantity aliases do not prefer Number To Nest.'
Assert ($com.Contains("[void](SN-Verify-WorkspaceCLData")) 'Import verification output is not suppressed.'
Assert ($worker.Contains('$rawEngineData=@(if($mode -eq ''IMPORT_ONLY'')')) 'Worker does not capture complete engine pipeline output.'
Assert ($worker.Contains('$engineResults=@($rawEngineData | Where-Object')) 'Worker does not filter for the structured engine result.'
Assert ($taskpane.Contains('function normalizeParts(parts)')) 'Task pane null-part guard missing.'
Assert ($taskpane.Contains("No valid CL parts are available for AutoTask")) 'Task pane AutoTask input guard missing.'

# Import must remain task-free; AutoTask owns task creation.
$importStart=$com.IndexOf('function Invoke-SigmaNestImportGeometry')
$autoStart=$com.IndexOf('function Invoke-SigmaNestAutoTask')
$importBody=$com.Substring($importStart,$autoStart-$importStart)
Assert (-not $importBody.Contains("CreateTasksListForNewPartsInWS()")) 'Import Geometry must not create TasksList entries.'
Assert (-not $importBody.Contains("AutoTask()")) 'Import Geometry must not run AutoTask.'

Write-Host ("SigmaNEST static release gate: PASS (bridge $serverVersion)") -ForegroundColor Green
