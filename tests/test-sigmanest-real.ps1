param(
  [Parameter(Mandatory=$true)][string]$WorkspacePath,
  [string]$OutputDirectory = $(Join-Path $env:TEMP 'CL-WS-Creator-SigmaNEST-SelfTest'),
  [switch]$KeepWorkspace
)
$ErrorActionPreference='Stop'
if(-not(Test-Path -LiteralPath $WorkspacePath)){throw "Workspace not found: $WorkspacePath"}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$source=[IO.Path]::GetFullPath($WorkspacePath)
$copy=Join-Path $OutputDirectory ("SELFTEST-"+[Guid]::NewGuid().ToString('N')+'.ws')
Copy-Item -LiteralPath $source -Destination $copy -Force

$app=$null
try{
  $app=New-Object -ComObject SigmaNEST.SNApp
  if($null -eq $app){throw 'SigmaNEST.SNApp could not be created.'}

  $app.LoadWorkSpaceFile([string]$copy)
  $before=0
  try{$before=[int]$app.PartsList.Count}catch{}
  if($before -lt 1){throw 'Self-test workspace contains no PartsList entries.'}

  Write-Host "PartsList before task creation: $before"
  $app.CreateTasksListForNewPartsInWS()
  Start-Sleep -Milliseconds 500
  $createdTasks=0
  try{$createdTasks=[int]$app.TasksList.Count}catch{}
  if($createdTasks -lt 1){throw 'CreateTasksListForNewPartsInWS created no TasksList entries.'}
  Write-Host "TasksList after task creation: $createdTasks" -ForegroundColor Green

  $app.AutoTask()
  Start-Sleep -Milliseconds 1500
  $after=0
  try{$after=[int]$app.TasksList.Count}catch{}
  if($after -lt 1){throw 'AutoTask left no TasksList entries.'}
  Write-Host "TasksList after AutoTask: $after" -ForegroundColor Green

  try{$app.RefreshTreeView()}catch{}
  try{$app.Redraw()}catch{}
  $result=[pscustomobject]@{
    ok=$true
    workspace=$copy
    parts=$before
    tasksBeforeAutoTask=$createdTasks
    tasksAfterAutoTask=$after
    sigmaNestCom='PASS'
    timestamp=(Get-Date).ToString('o')
  }
  $result | ConvertTo-Json -Depth 10
}finally{
  if($app){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{}}
  if(-not $KeepWorkspace -and (Test-Path -LiteralPath $copy)){
    Remove-Item -LiteralPath $copy -Force -ErrorAction SilentlyContinue
  }
}
