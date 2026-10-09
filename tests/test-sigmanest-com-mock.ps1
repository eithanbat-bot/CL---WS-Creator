# Deterministic SigmaNEST COM simulation for the WS-first workflow.
# Does not require SigmaNEST; it verifies that import writes all production
# values and the tasking stage consumes only the saved workspace snapshot.

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

# --- Import Geometry persists CL data onto the part ---
$part=[pscustomobject]@{
  Name='TEST-001'
  DrawingNumber=''
  PartFilename='TEST-001.DXF'
  QtyToNest=99
  NumberToNest=99
  BatchQty=1
  Material='OLD'
  Thickness=3.0
}
$app=[pscustomobject]@{PartsList=(New-List @($part))}
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
}

$updated=@(SN-Apply-WorkspacePartData -app $app -requestParts @($requestPart) -jobName 'SELFTEST')
Assert ($updated.Count -eq 1) 'Part-data update did not return exactly one record.'
Assert ($part.QtyToNest -eq 4) ("QtyToNest was not set to the per-vehicle CL quantity 4; got $($part.QtyToNest).")
Assert ($part.BatchQty -eq 4) ("BatchQty was not set to the CL multiplier 4; got $($part.BatchQty).")
Assert ($part.Material -eq 'MS') ("Material was not changed to the mapped CL material MS; got $($part.Material).")
Assert ([double]$part.Thickness -eq 8.0) ("Thickness was not changed to 8mm; got $($part.Thickness).")
Assert ([int]$updated[0].qty -eq 4) ("Reported per-vehicle quantity was not 4; got $($updated[0].qty).")
Assert ([int]$updated[0].batchMultiplier -eq 4) ("Reported batch multiplier was not 4; got $($updated[0].batchMultiplier).")
Assert ($updated[0].batchProperty) 'Import did not record the verified BatchQty property.'

[void](SN-Verify-WorkspaceCLData -app $app -requestParts @($requestPart))
$snapshot=@(SN-Read-WorkspaceProductionData -app $app)
Assert ($snapshot.Count -eq 1) 'Saved workspace snapshot did not contain exactly one part.'
Assert ($snapshot[0].sigmaMaterial -eq 'MS') 'Workspace snapshot did not read material from the part.'
Assert ([double]$snapshot[0].thicknessMm -eq 8.0) 'Workspace snapshot did not read thickness from the part.'
Assert ([int]$snapshot[0].qty -eq 4) 'Workspace snapshot did not read per-vehicle quantity from the part.'
Assert ([int]$snapshot[0].batchMultiplier -eq 4) 'Workspace snapshot did not read the BatchQty multiplier from the part.'

# --- AutoTask labeling/batching uses only the saved-WS snapshot ---
$taskPart=[pscustomobject]@{
  Name='TEST-001'
  DrawingNumber='TEST-001'
  PartFilename='TEST-001.DXF'
  Material='MS'
  Thickness=8.0
  QtyToNest=4
  BatchQty=4
}
$task=[pscustomobject]@{PartsList=(New-List @($taskPart)); Name='OLD TASK'; BatchMultiplier=1}
$taskApp=[pscustomobject]@{TasksList=(New-List @($task))}
$taskResult=SN-Set-TaskNameAndBatch -app $taskApp -workspaceParts $snapshot

Assert ($taskResult.tasks.Count -eq 1) 'Task result did not contain exactly one task.'
Assert ($taskResult.tasks[0].labelApplied) 'Task label was not applied in the COM simulation.'
Assert ($taskResult.tasks[0].batchApplied) 'Task batch was not applied in the COM simulation.'
Assert ([int]$taskResult.tasks[0].batchMultiplier -eq 4) 'Task batch multiplier was not read from the WS snapshot.'
Assert ($task.Name -eq '01 | MS | 8mm') ("Task label is wrong: $($task.Name).")
Assert ([int]$task.BatchMultiplier -eq 4) 'Task BatchMultiplier was not set to the saved WS batch multiplier.'
Assert ([int]$taskPart.QtyToNest -eq 4) 'Tasking changed the per-vehicle quantity from the saved WS.'
Assert ([int]$taskPart.BatchQty -eq 4) 'Tasking changed the persisted BatchQty on the saved WS part.'

Write-Host 'SigmaNEST WS-first mock integration: PASS' -ForegroundColor Green
