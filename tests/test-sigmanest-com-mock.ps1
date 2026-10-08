# Deterministic SigmaNEST COM simulation.
# This does not require SigmaNEST. It exercises the same task/part mutation
# functions against COM-shaped PowerShell objects so regressions are caught
# before a release.

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

# --- Part-data quantity contract ---
$part=[pscustomobject]@{
  Name='TEST-001'
  DrawingNumber=''
  PartFilename='TEST-001.DXF'
  QtyToNest=16
  NumberToNest=16
  BatchQty=16
  Material='OLD'
  Thickness=3.0
}
$partsList=New-List @($part)
$app=[pscustomobject]@{PartsList=$partsList}
$requestPart=[pscustomobject]@{
  part='TEST-001'
  sourcePath='C:\TEST\TEST-001.DXF'
  qty=4
  sigmaMaterial='MS'
  thicknessMm=8
}
$updated=@(SN-Apply-WorkspacePartData -app $app -requestParts @($requestPart) -jobName 'SELFTEST')
Assert ($updated.Count -eq 1) 'Part-data update did not return exactly one record.'
Assert ($part.QtyToNest -eq 4) ("QtyToNest was not changed to 4; got $($part.QtyToNest).")
Assert ($part.BatchQty -eq 16) ("BatchQty was incorrectly changed during part import; got $($part.BatchQty).")
Assert ($part.Material -eq 'MS') ("Material was not changed to MS; got $($part.Material).")
Assert ([double]$part.Thickness -eq 8.0) ("Thickness was not changed to 8mm; got $($part.Thickness).")
Assert ([int]$updated[0].qty -eq 4) ("Reported CL quantity was not 4; got $($updated[0].qty).")

# --- Task material/thickness/quantity/label/batch contract ---
$taskPart=[pscustomobject]@{
  Name='TEST-001'
  Material='OLD'
  Thickness=3.0
  QtyToNest=99
  NumberToNest=99
  BatchQty=99
  Quantity=99
  Qty=99
}
$task=[pscustomobject]@{PartsList=(New-List @($taskPart)); Name='OLD TASK'; BatchMultiplier=1}
$taskList=New-List @($task)
$taskApp=[pscustomobject]@{TasksList=$taskList}

$taskRequest=[pscustomobject]@{
  part='TEST-001'
  qty=4
  sigmaMaterial='MS'
  thicknessMm=8.0
  batchMultiplier=2
  taskBatches=@()
}

SN-Set-TaskMaterialAndThickness -app $taskApp -requestParts @($taskRequest)
SN-Set-TaskPartQuantity -app $taskApp -requestParts @($taskRequest)
$taskResult=SN-Set-TaskNameAndBatch -app $taskApp -requestParts @($taskRequest)

Assert ($taskPart.Material -eq 'MS') ("Task material was not MS; got $($taskPart.Material).")
Assert ([double]$taskPart.Thickness -eq 8.0) ("Task thickness was not 8mm; got $($taskPart.Thickness).")
Assert ([double]$taskPart.QtyToNest -eq 4) ("Task quantity was not 4; got $($taskPart.QtyToNest).")
Assert ([double]$taskPart.BatchQty -eq 99) ("Task quantity operation unexpectedly changed part BatchQty; got $($taskPart.BatchQty).")
Assert ($taskResult.tasks.Count -eq 1) 'Task result did not contain exactly one task.'
Assert ($taskResult.tasks[0].labelApplied) 'Task label was not applied in the COM simulation.'
Assert ($taskResult.tasks[0].batchApplied) 'Task batch was not applied in the COM simulation.'
Assert ([int]$taskResult.tasks[0].batchMultiplier -eq 2) 'Task batch multiplier was not 2.'
Assert ($task.Name -eq '01 | MS | 8mm') ("Task label is wrong: $($task.Name).")
Assert ([int]$task.BatchMultiplier -eq 2) 'Task BatchMultiplier was not 2.'

Write-Host 'SigmaNEST mock integration: PASS' -ForegroundColor Green
