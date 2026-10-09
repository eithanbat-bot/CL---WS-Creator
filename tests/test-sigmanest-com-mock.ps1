# Deterministic SigmaNEST COM simulation for Import Geometry + WS-only AutoTask.
# Verifies PartList values are imported per vehicle, while batch multiplier data
# is staged on TasksList so SigmaNEST cannot multiply QtyToNest twice.

$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
. (Join-Path $Root 'bridge\sigmanest-com.ps1')

function Assert([bool]$condition,[string]$message){
  if(-not $condition){throw $message}
}
function New-List($items){
  $list=[pscustomobject]@{Count=@($items).Count;_items=@($items)}
  $list | Add-Member -MemberType ScriptMethod -Name Items -Value { param($i) return $this._items[[int]$i] }
  return $list
}

$requestPart=[pscustomobject]@{
  part='TEST-001'
  description='TEST BRACKET'
  sourcePath='C:\TEST\TEST-001.DXF'
  sourceType='DXF'
  qty=4
  effectiveQty=16
  batchMultiplier=4
  sigmaMaterial='MS'
  thicknessMm=8
  material='8mm Mild Steel'
  taskBatches=@(@{sheet='BAT - Test';batchMultiplier=4;vehicleQty=4})
  sourceSheets=@('BAT - Test')
  sourceRows=@(2)
}

# --- Import Geometry: PartList production fields, BatchQty staged at x1 ---
$part=[pscustomobject]@{
  Name='TEST-001'
  DrawingNumber=''
  PartFilename='C:\TEST\TEST-001.DXF'
  QtyToNest=99
  NumberToNest=99
  BatchQty=1
  Material='OLD'
  Thickness=3.0
}
$partsApp=[pscustomobject]@{PartsList=(New-List @($part))}
$partUpdates=@(SN-Apply-WorkspacePartData -app $partsApp -requestParts @($requestPart) -jobName 'SELFTEST' -DeferBatchToTasks)
Assert ($partUpdates.Count -eq 1) 'Import part update did not return exactly one record.'
Assert ([int]$part.QtyToNest -eq 4) ("PartList QtyToNest must remain per-vehicle 4; got $($part.QtyToNest).")
Assert ([int]$part.BatchQty -eq 1) ("PartList BatchQty must remain x1 until task list creation; got $($part.BatchQty).")
Assert ($part.Material -eq 'MS') ("PartList material was not changed to MS; got $($part.Material).")
Assert ([double]$part.Thickness -eq 8.0) ("PartList thickness was not changed to 8mm; got $($part.Thickness).")
Assert ([int]$partUpdates[0].batchMultiplier -eq 4) 'Import result did not preserve the CL batch multiplier.'
Assert ([string]$partUpdates[0].batchProperty -eq 'TASKS_LIST.PartsList.BatchQty') 'Import result did not identify where batch data will be persisted.'

[void](SN-Verify-WorkspaceCLData -app $partsApp -requestParts @($requestPart))

# --- Import Geometry prepares TasksList production data, but does not AutoTask ---
$taskPart=[pscustomobject]@{
  Name='TEST-001'
  DrawingNumber='TEST-001'
  PartFilename='C:\TEST\TEST-001.DXF'
  Material='OLD'
  Thickness=3.0
  QtyToNest=99
  BatchQty=1
}
$task=[pscustomobject]@{
  PartsList=(New-List @($taskPart))
  Name='OLD TASK'
  BatchQuantity=1
  BatchMultiplier=1
}
$taskApp=[pscustomobject]@{TasksList=(New-List @($task))}
$taskUpdates=@(SN-Apply-TaskCLData -app $taskApp -requestParts @($requestPart))
Assert ($taskUpdates.Count -eq 1) 'CL task-data preparation did not return exactly one verified task-part update.'
Assert ([int]$task.BatchQuantity -eq 4) ("TasksList BatchQuantity was not set to 4; got $($task.BatchQuantity).")
Assert ($taskPart.Material -eq 'MS') ("TasksList part material was not changed to MS; got $($taskPart.Material).")
Assert ([double]$taskPart.Thickness -eq 8.0) ("TasksList part thickness was not changed to 8mm; got $($taskPart.Thickness).")
Assert ([int]$taskPart.BatchQty -eq 4) ("TasksList part BatchQty was not set to 4; got $($taskPart.BatchQty).")
Assert ([int]$taskPart.QtyToNest -eq 4) ("TasksList QtyToNest was multiplied; expected per-vehicle 4, got $($taskPart.QtyToNest).")
Assert ([string]$taskUpdates[0].batchProperty -eq 'TASKS_LIST.PartsList.BatchQty') 'Task data update did not identify persisted BatchQty.'

[void](SN-Verify-WorkspaceTaskData -app $taskApp -requestParts @($requestPart))
$snapshot=@(SN-Read-WorkspaceProductionData -app $taskApp)
Assert ($snapshot.Count -eq 1) 'Saved-WS snapshot did not contain exactly one TasksList part.'
Assert ($snapshot[0].sigmaMaterial -eq 'MS') 'Saved-WS snapshot did not read material from TasksList.'
Assert ([double]$snapshot[0].thicknessMm -eq 8.0) 'Saved-WS snapshot did not read thickness from TasksList.'
Assert ([int]$snapshot[0].qty -eq 4) 'Saved-WS snapshot did not read per-vehicle quantity from TasksList.'
Assert ([int]$snapshot[0].batchMultiplier -eq 4) 'Saved-WS snapshot did not read the batch multiplier from TasksList.'

# --- Tasking uses only the saved snapshot: no external CL payload is required ---
$taskResult=SN-Set-TaskNameAndBatch -app $taskApp -workspaceParts $snapshot
Assert ($taskResult.tasks.Count -eq 1) 'Task result did not contain exactly one task.'
Assert ($taskResult.tasks[0].labelApplied) 'Task label was not applied in the COM simulation.'
Assert ($taskResult.tasks[0].batchApplied) 'Task batch was not applied in the COM simulation.'
Assert ([int]$taskResult.tasks[0].batchMultiplier -eq 4) 'Task batch multiplier was not read from the saved-WS snapshot.'
Assert ($task.Name -eq '01 | MS | 8mm') ("Task label is wrong: $($task.Name).")
Assert ([int]$task.BatchQuantity -eq 4) 'Task batch was not kept at the saved-WS multiplier.'
Assert ([int]$taskPart.QtyToNest -eq 4) 'Task labeling changed the per-vehicle quantity.'
Assert ([int]$taskPart.BatchQty -eq 4) 'Task labeling changed the saved TasksList BatchQty.'

Write-Host 'SigmaNEST Import + WS-only AutoTask mock integration: PASS' -ForegroundColor Green
