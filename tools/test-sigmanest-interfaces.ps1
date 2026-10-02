$ErrorActionPreference='Continue'
Write-Host '=== SigmaNEST NESTED COM INTERFACE DIAGNOSTIC ===' -ForegroundColor Cyan
Write-Host 'Keep SigmaNEST open.'
Write-Host ''
function Show-ObjectMembers($label,$obj){
  Write-Host ('['+$label+']') -ForegroundColor Yellow
  if($null -eq $obj){Write-Host '  OBJECT: <null>' -ForegroundColor Red;Write-Host '';return}
  try {
    $obj | Get-Member -MemberType Methods,Properties -ErrorAction Stop | ForEach-Object { Write-Host ('  '+$_.Name+' :: '+$_.Definition) }
  } catch { Write-Host ('  Get-Member error: '+$_.Exception.Message) -ForegroundColor Red }
  Write-Host ''
}
try {
  $app=New-Object -ComObject SigmaNEST.SNApp
  Show-ObjectMembers 'SNApp.PartsLibrary' $app.PartsLibrary
  Show-ObjectMembers 'SNApp.PartsList' $app.PartsList
  Show-ObjectMembers 'SNApp.WorkOrdersList' $app.WorkOrdersList
  Show-ObjectMembers 'SNApp.TasksList' $app.TasksList
} catch { Write-Host ('SNApp create failed: '+$_.Exception.Message) -ForegroundColor Red }
try {
  $auto=New-Object -ComObject SigmaNEST.SNAutomation
  Show-ObjectMembers 'SNAutomation.PartList' $auto.PartList
  Show-ObjectMembers 'SNAutomation.WOList' $auto.WOList
} catch { Write-Host ('SNAutomation create failed: '+$_.Exception.Message) -ForegroundColor Red }
Write-Host '=== END ===' -ForegroundColor Cyan
Write-Host 'Copy all output and send it back.'