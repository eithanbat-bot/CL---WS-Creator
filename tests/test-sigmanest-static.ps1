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
Assert ($server.Contains("workerVersion='$serverVersion'")) 'Initial worker status does not report the current bridge version.'
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
Assert ($serverVersion -eq '2.26.0') "Unexpected bridge version: $serverVersion"

Assert ($com.Contains("function Invoke-SigmaNestImportGeometry")) 'Import entry point missing.'
Assert ($com.Contains("function Invoke-SigmaNestAutoTask")) 'AutoTask entry point missing.'
Assert ($com.Contains("CreateTasksListForNewPartsInWS")) 'Task creation method missing.'
Assert ($com -match '\$phase\s*=\s*''CREATE_TASKS_FOR_IMPORTED_PARTS''') 'AutoTask phase does not explicitly create tasks before AutoTask.'
Assert ($com -match '\$phase\s*=\s*''AUTO_TASK''') 'AutoTask phase missing.'
Assert ($com.Contains('try{[void]$app.RefreshTreeView()}catch{}')) 'SigmaNEST tree refresh missing or return value is not suppressed.'
Assert ($com.Contains('function SN-Set-QtyToNest')) 'Canonical SigmaNEST QtyToNest setter is missing.'
Assert ($com.Contains("'NumberToNest','NumberToLoad','QtyToNest")) 'Fallback part quantity aliases are not present.'
Assert ($com.Contains("[void](SN-Verify-WorkspaceCLData")) 'Import verification output is not suppressed.'
Assert ($worker.Contains('$rawEngineData=@(if($mode -eq ''IMPORT_ONLY'')')) 'Worker does not capture complete engine pipeline output.'
Assert ($worker.Contains('$engineResults=@($rawEngineData | Where-Object')) 'Worker does not filter for the structured engine result.'
Assert ($worker.Contains('$rawAutoTaskData=@(Invoke-SigmaNestAutoTask')) 'Worker does not capture AutoTask pipeline output.'
Assert ($worker.Contains('$autoTaskResults=@($rawAutoTaskData | Where-Object')) 'Worker does not filter AutoTask for the structured result.'
Assert ($com.Contains('[void]$app.CreatePartsListForNewPartsInWS()')) 'Import geometry return value is not suppressed.'
Assert ($com.Contains('[void]$automation.FileNew()')) 'Import Geometry does not initialize a new SigmaNEST workspace.'
Assert ($com.Contains('[void]$automation.Add2DDXFPart([string]$sourcePath)')) 'Import Geometry does not use the tested DXF import method.'
Assert ($com.Contains('[void]$automation.FileSave([string]$wsPath,0,0)')) 'Import Geometry does not use the tested workspace save method.'
Assert ($updater.Contains('$entry.ValidatedStage')) 'Updater does not install from the validated runtime stage.'
Assert ($updater.Contains('Stage=$manifestStage')) 'Updater does not install manifest from its validated staging copy.'
Assert ($com.Contains("function SN-Queue-Geometry(`$app,[string]`$sourcePath,[string]`$sourceType,`$clData=`$null,`$automation=`$null)")) 'Geometry queue does not accept the workspace automation object.'
Assert ($com.Contains('[void]$app.CreateTasksListForNewPartsInWS()')) 'Task creation pipeline output is not suppressed.'
Assert ($com.Contains('[void]$app.AutoTask()')) 'AutoTask pipeline output is not suppressed.'
Assert ($taskpane.Contains('function normalizeParts(parts)')) 'Task pane null-part guard missing.'
Assert ($taskpane.Contains("No valid CL parts are available for AutoTask")) 'Task pane AutoTask input guard missing.'

# Import must remain task-free; AutoTask owns task creation.
$importStart=$com.IndexOf('function Invoke-SigmaNestImportGeometry')
$autoStart=$com.IndexOf('function Invoke-SigmaNestAutoTask')
$importBody=$com.Substring($importStart,$autoStart-$importStart)
Assert (-not $importBody.Contains("CreateTasksListForNewPartsInWS()")) 'Import Geometry must not create TasksList entries.'
Assert (-not $importBody.Contains("AutoTask()")) 'Import Geometry must not run AutoTask.'
Assert ($importBody.Contains('[void]$app.CreatePartsListForNewPartsInWS()')) 'Import Geometry must suppress the PartsList-creation return value.'

Write-Host ("SigmaNEST static release gate: PASS (bridge $serverVersion)") -ForegroundColor Green
