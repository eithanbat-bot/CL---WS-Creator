$ErrorActionPreference = "Continue"
Write-Host "=== SigmaNEST COM METHOD SIGNATURES ===" -ForegroundColor Cyan
Write-Host "Keep SigmaNEST open.`n"

$tests = @(
  @{ ProgId="SigmaNEST.SNAutomation"; Members=@("LoadWSFile","LoadWOLFile","PartList","WOList","FileNew","FileSave","RunScript","ExeBatchString2") },
  @{ ProgId="SigmaNEST.SNApp"; Members=@("LoadPart","LoadSNParts","ImportWorkspace","LoadWorkSpaceFile","SaveWorkSpaceFile","CreatePartsListForNewPartsInWS","CreateTasksListForNewPartsInWS","RemoveNewPartsFromWorkspace","RemovePartFromWsPartList","UpdateAllPartsInTask","ExecuteBatchCommand","ExecuteBatchFile","ExecuteBatchString","PartsLibrary","PartsList","WorkOrdersList","TasksList","Version") },
  @{ ProgId="SigmaNEST.SNWorkOrder"; Members=@("AddWO","AddPartToWO","ModifyWO","DeleteWO","LoadAllWOParts","CheckDuplicateWO") },
  @{ ProgId="SigmaNEST.SNTask"; Members=@("CreateTask","ImportTask","ImportTask2","SetMachine","SetNestSheet","SetNestParam","SetNCProgramName","SetActiveTask","ExportTask") },
  @{ ProgId="SigmaNEST.SNPartExportImport"; Members=@("ImportPart","ImportPart2","ImportPartWithFeedback","ImportAsParts","PartImportSettings","ExportPart") },
  @{ ProgId="SigmaNEST.SNPaths"; Members=@("GetPath","SetPath") }
)

foreach($t in $tests){
  Write-Host ("["+$t.ProgId+"]") -ForegroundColor Yellow
  try{$obj=New-Object -ComObject $t.ProgId -ErrorAction Stop}catch{Write-Host ("  CREATE FAILED: "+$_.Exception.Message) -ForegroundColor Red;continue}
  foreach($name in $t.Members){
    try{
      $g=$obj | Get-Member -Name $name -ErrorAction Stop
      foreach($m in @($g)){Write-Host ("  "+$m.Name+" :: "+$m.Definition)}
    }catch{Write-Host ("  "+$name+" :: NOT EXPOSED") -ForegroundColor DarkGray}
  }
  Write-Host ""
}

Write-Host "=== SAFE PROPERTIES ===" -ForegroundColor Cyan
try{$app=New-Object -ComObject SigmaNEST.SNApp
  foreach($p in @("Version","Name","MetricUnits","DatabaseKind","ADOConnectionString","DocMode","UIMode")){
    try{Write-Host ("SNApp."+($p)+" = "+$app.$p)}catch{Write-Host ("SNApp."+($p)+" = <unavailable>")}
  }
}catch{}
Write-Host ""
Write-Host "=== END ===" -ForegroundColor Cyan
Write-Host "Copy all output and send it back to ChatGPT."