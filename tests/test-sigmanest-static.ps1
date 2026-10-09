# Static release-gate checks for the SigmaNEST adapter.
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
$server=Get-Content -Raw (Join-Path $Root 'bridge\server.ps1')
$worker=Get-Content -Raw (Join-Path $Root 'bridge\build-worker.ps1')
$updater=Get-Content -Raw (Join-Path $Root 'bridge\update-and-start.ps1')
$com=Get-Content -Raw (Join-Path $Root 'bridge\sigmanest-com.ps1')
$taskpane=Get-Content -Raw (Join-Path $Root 'taskpane.js')

function Assert([bool]$condition,[string]$message){if(-not $condition){throw $message}}

$serverVersion=[regex]::Match($server,'(?m)^\s*\$BRIDGE_VERSION\s*=\s*''([^'']+)''').Groups[1].Value
$workerVersion=[regex]::Match($worker,'(?m)^\s*\$WorkerVersion\s*=\s*''([^'']+)''').Groups[1].Value
Assert ($serverVersion -eq $workerVersion) "Bridge/worker version mismatch: $serverVersion vs $workerVersion"
Assert ($server.Contains('workerVersion=$BRIDGE_VERSION')) 'Initial worker status does not report the current bridge version.'
$matcherStart=$server.IndexOf('function Find-Part(')
$matcherEnd=$server.IndexOf('function Csv(', $matcherStart)
Assert ($matcherStart -ge 0 -and $matcherEnd -gt $matcherStart) 'Geometry matcher entry point is missing.'
$matcher=$server.Substring($matcherStart,$matcherEnd-$matcherStart)
Assert ($matcher.Contains('$exactCandidates=if($dxfExact.Count -gt 0){$dxfExact}else{$prsExact}')) 'Exact matching can let PRS override an available DXF.'
Assert ($matcher.Contains('$embeddedCandidates=if($dxfEmbedded.Count -gt 0){$dxfEmbedded}else{$prsEmbedded}')) 'Embedded matching can let PRS override an available DXF.'
Assert ($matcher.Contains('$variationCandidates=if($dxfVars.Count -gt 0){$dxfVars}else{$prsVars}')) 'Variation matching can let PRS override an available DXF.'
$thicknessStart=$server.IndexOf('function Thickness-Number(')
$thicknessEnd=$server.IndexOf('function Material-Equal(', $thicknessStart)
Assert ($thicknessStart -ge 0 -and $thicknessEnd -gt $thicknessStart) 'Thickness parser is missing.'
$thicknessParser=$server.Substring($thicknessStart,$thicknessEnd-$thicknessStart)
Assert ($thicknessParser.Contains("'^\s*(\d+(?:\.\d+)?)\s*$'")) 'Thickness parser does not accept plain numeric CL thickness cells.'
$shardStart=$server.IndexOf('function Get-DxfShardCandidates(')
$shardEnd=$server.IndexOf('function Is-EmbeddedMatch', $shardStart)
Assert ($shardStart -ge 0 -and $shardEnd -gt $shardStart) 'DXF shard candidate reader is missing.'
$shardReader=$server.Substring($shardStart,$shardEnd-$shardStart)
Assert ($shardReader.Contains('$columns=$line.Split([char]9)')) 'DXF shard reader does not split TSV columns correctly.'
Assert ($shardReader.Contains('$file=[string]$columns[1]')) 'DXF shard reader does not use the File column.'
Assert (-not $shardReader.Contains('$file=$line.Substring($tab+1)')) 'DXF shard reader still appends timestamp/length to the file path.'
Assert ($serverVersion -eq '2.32.2') "Unexpected bridge version: $serverVersion"

Assert ($com.Contains("function Invoke-SigmaNestImportGeometry")) 'Import entry point missing.'
Assert ($com.Contains("function Invoke-SigmaNestAutoTask")) 'AutoTask entry point missing.'
Assert ($com.Contains("CreateTasksListForNewPartsInWS")) 'Task creation method missing.'
Assert ($com -match '\$phase\s*=\s*''CREATE_TASKS_FOR_IMPORTED_PARTS''') 'AutoTask phase does not explicitly create tasks before AutoTask.'
Assert ($com -match '\$phase\s*=\s*''AUTO_TASK''') 'AutoTask phase missing.'
Assert ($com.Contains('try{[void]$app.RefreshTreeView()}catch{}')) 'SigmaNEST tree refresh missing or return value is not suppressed.'
Assert ($com.Contains('function SN-Set-QtyToNest')) 'Canonical SigmaNEST QtyToNest setter is missing.'
Assert ($com.Contains("'BatchQty','BatchQuantity','BatchMultiplier','Batch'")) 'Import does not write a persisted BatchQty multiplier to the SigmaNEST PartsList.'
Assert ($com.Contains('BatchQty multiplier is not')) 'Import verification does not verify the persisted batch multiplier.'
Assert ($com.Contains('function SN-Read-WorkspaceProductionData($app)')) 'AutoTask is missing the saved-WS production-data reader.'
Assert ($com.Contains('effectiveBatchQuantity=')) 'CL link file does not retain effective batch quantities.'
Assert ($com.Contains("'NumberToNest','NumberToLoad','QtyToNest")) 'Fallback part quantity aliases are not present.'
Assert ($com.Contains("[void](SN-Verify-WorkspaceCLData")) 'Import verification output is not suppressed.'
Assert ($worker.Contains('$rawEngineData=@(if($mode -eq ''IMPORT_ONLY'')')) 'Worker does not capture complete engine pipeline output.'
Assert ($worker.Contains('$engineResults=@($rawEngineData | Where-Object')) 'Worker does not filter for the structured engine result.'
Assert ($worker.Contains('$rawAutoTaskData=@(Invoke-SigmaNestAutoTask')) 'Worker does not capture AutoTask pipeline output.'
Assert ($worker.Contains('$autoTaskResults=@($rawAutoTaskData | Where-Object')) 'Worker does not filter AutoTask for the structured result.'
Assert ($com.Contains('[void]$app.CreatePartsListForNewPartsInWS()')) 'Import geometry return value is not suppressed.'
Assert ($com.Contains('[void]$automation.FileNew()')) 'Import Geometry does not initialize a new SigmaNEST workspace.'
Assert ($com.Contains("method='LoadPart-DXF-GEOMETRY'")) 'DXF import method result is not reported.'
$dxfQueueStart=$com.IndexOf("if(`$sourceType -eq 'DXF')")
$dxfQueueEnd=$com.IndexOf('# PRS is an explicit fallback', $dxfQueueStart)
Assert ($dxfQueueStart -ge 0 -and $dxfQueueEnd -gt $dxfQueueStart) 'DXF queue branch is missing.'
$dxfQueue=$com.Substring($dxfQueueStart,$dxfQueueEnd-$dxfQueueStart)
Assert ($dxfQueue.Contains('[void]$app.LoadPart([string]$sourcePath)')) 'DXF import does not load the exact DXF into a fresh workspace.'
Assert (-not $dxfQueue.Contains('$automation.Add2DDXFPart(')) 'DXF import still calls Add2DDXFPart, which does not add parts on an empty workspace.'
Assert ($com.Contains('[void]$automation.FileSave([string]$wsPath,0,0)')) 'Import Geometry does not use the tested workspace save method.'
Assert ($updater.Contains('$entry.ValidatedStage')) 'Updater does not install from the validated runtime stage.'
Assert ($updater.Contains('Stage=$manifestStage')) 'Updater does not install manifest from its validated staging copy.'
Assert ($updater.Contains("$stageRoot=Join-Path `$BridgeDir ('_update-stage-'+[Guid]::NewGuid().ToString('N'))")) 'Updater validated staging is not isolated from TEMP cleanup.'
Assert ($updater.Contains('?cb=$cacheBust')) 'Updater does not cache-bust its refreshed bootstrap download.'
$bat=Get-Content -Raw (Join-Path $Root 'Start Bridge Fixed.bat')
Assert ($bat.Contains('?cb=') -and $bat.Contains('$cb')) 'Bootstrap does not cache-bust the downloaded updater.'
Assert ($com.Contains("function SN-Queue-Geometry(`$app,[string]`$sourcePath,[string]`$sourceType,`$clData=`$null,`$automation=`$null)")) 'Geometry queue does not accept the workspace automation object.'
Assert ($com.Contains('[void]$app.CreateTasksListForNewPartsInWS()')) 'Task creation pipeline output is not suppressed.'
Assert ($com.Contains('[void]$app.AutoTask()')) 'AutoTask pipeline output is not suppressed.'
Assert ($com.Contains('$multis+=[int]$rp.batchMultiplier')) 'Task batch logic does not read multipliers from the saved workspace snapshot.'
Assert ($com.Contains('[void]$automation.AddPartImport([string]$x.sourcePath,[double]1.0,[double]1.0,0,0,0)')) 'DXF import does not use SigmaNEST AddPartImport.'
Assert ($com.Contains("importMethod=`$(if(`$sourceType -eq 'DXF'){'AddPartImport-DXF-GEOMETRY'}else{'LoadPart-PRS-GEOMETRY'})")) 'Import result does not distinguish DXF and PRS loading.'
Assert ($com.Contains('[void]$automation.ResetSigmaNEST()')) 'Import/AutoTask does not clear stale SigmaNEST workspace rows before loading.'
Assert ($com.Contains("SN-Save-WorkspaceVerified -app `$app -wsPath `$wsPath -label 'CL-data PartsList and TasksList import'")) 'CL-data import does not save the populated SNApp workspace.'
$autoBodyStart=$com.IndexOf('function Invoke-SigmaNestAutoTask(')
$autoBodyEnd=$com.IndexOf('function Invoke-SigmaNestBuild(', $autoBodyStart)
Assert ($autoBodyStart -ge 0 -and $autoBodyEnd -gt $autoBodyStart) 'AutoTask method body is missing.'
$autoBody=$com.Substring($autoBodyStart,$autoBodyEnd-$autoBodyStart)
Assert ($autoBody.Contains('[void]$automation.FileNew()')) 'AutoTask does not reset the existing SigmaNEST workspace before loading the saved job.'
Assert ($autoBody.IndexOf('[void]$automation.FileNew()') -lt $autoBody.IndexOf('[void]$app.LoadWorkSpaceFile([string]$wsPath)')) 'AutoTask loads the saved job before clearing existing PartsList entries.'
Assert ($autoBody.Contains('SN-Read-WorkspaceProductionData -app $app')) 'AutoTask does not snapshot production values from the saved WS.'
Assert ($autoBody.Contains('SN-Set-TaskNameAndBatch -app $app -workspaceParts $workspaceParts')) 'AutoTask does not use batch multipliers read from the saved WS.'
Assert (-not $autoBody.Contains('SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts')) 'AutoTask is still overwriting saved WS values from an Excel/CL request.'
Assert (-not $autoBody.Contains('CreateTasksListForNewPartsInWS()')) 'AutoTask must use the saved TasksList and must not recreate it.'
Assert (-not $autoBody.Contains('SN-Apply-WorkspacePartData -app $app -requestParts $Request.parts')) 'AutoTask must never reapply CL payload to PartsList.'
Assert ($autoBody.Contains('SN-Apply-TaskCLData -app $app -requestParts $workspaceParts')) 'AutoTask must restore/verify only production data read from the saved WS after AutoTask.'
Assert ($taskpane.Contains('function normalizeParts(parts)')) 'Task pane null-part guard missing.'
Assert ($taskpane.Contains('sole source of material, thickness, quantity and batch multiplier')) 'AutoTask task pane must explain that the saved WS is its sole production-data source.'
Assert (-not $taskpane.Substring($taskpane.IndexOf('async function autoTaskOrder()'),$taskpane.IndexOf('async function checkBridge(')-$taskpane.IndexOf('async function autoTaskOrder()')).Contains('readCL()')) 'AutoTask must not reread or resend CL data.'
Assert ($taskpane.Contains("autoTaskOrder').disabled=!currentWsPath")) 'AutoTask must remain available from a saved WS without a loaded CL.'

# Import must prepare and verify CL production data on both PartsList and TasksList; AutoTask only runs later.
$importStart=$com.IndexOf('function Invoke-SigmaNestImportGeometry')
$autoStart=$com.IndexOf('function Invoke-SigmaNestAutoTask')
$importBody=$com.Substring($importStart,$autoStart-$importStart)
Assert ($importBody.Contains('$automation.AddPartImport([string]$x.sourcePath')) 'Import Geometry does not use the tested DXF AddPartImport method.'
Assert ($importBody.Contains("importMethod=`$(if(`$sourceType -eq 'DXF'){'AddPartImport-DXF-GEOMETRY'}else{'LoadPart-PRS-GEOMETRY'})")) 'Import Geometry reports incorrect geometry source methods.'
Assert ($importBody.Contains('$app.LoadPart([string]$x.sourcePath)')) 'Import Geometry does not load explicit PRS sources.'
Assert (-not $importBody.Contains('LoadPart-DXF-GEOMETRY')) 'Import Geometry still reports the unverified DXF LoadPart route.'
Assert ($importBody.Contains('CreateTasksListForNewPartsInWS()')) 'Import Geometry must prepare a TasksList to persist CL batch metadata before AutoTask.'
Assert (-not $importBody.Contains('AutoTask()')) 'Import Geometry must prepare the workspace but must not run AutoTask.'
Assert ($importBody.Contains('-DeferBatchToTasks')) 'Import Geometry must keep the PartsList BatchQty at x1 until TasksList creation.'
Assert ($importBody.Contains('SN-Apply-TaskCLData -app $app -requestParts $queued')) 'Import Geometry must apply the batch multiplier and CL data to TasksList before saving.'
Assert ($importBody.Contains('SN-Verify-WorkspaceTaskData -app $app -requestParts $queued')) 'Import Geometry must verify the saved task-side CL data before enabling AutoTask.'
Assert ($importBody.Contains("[void](SN-Verify-WorkspaceCLData -app `$app -requestParts `$queued)")) 'Import Geometry must verify CL fields in the reopened saved PartsList.'
Assert ($importBody.Contains("[void](SN-Verify-WorkspaceTaskData -app `$app -requestParts `$queued)")) 'Import Geometry must verify batch and production fields in the reopened saved TasksList.'
Assert ($importBody.Contains("$phase='VERIFY_SAVED_CL_TASK_DATA'")) 'Import Geometry must verify the saved TasksList before releasing the workspace.'

Write-Host ("SigmaNEST static release gate: PASS (bridge $serverVersion)") -ForegroundColor Green
