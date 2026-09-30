$ErrorActionPreference='Continue'
Write-Host '=== SigmaNEST WORKSPACE / PART OBJECT DIAGNOSTIC ===' -ForegroundColor Cyan
Write-Host 'SigmaNEST must be open.'
Write-Host ''
$ws=Read-Host 'Enter the full path to a real SigmaNEST .ws file (example: U:\Jobs\JO 32 - 26 WS.xls)'
if([string]::IsNullOrWhiteSpace($ws)){Write-Host 'No workspace path supplied.' -ForegroundColor Red;exit}
if(-not(Test-Path -LiteralPath $ws)){Write-Host ('File not found: '+$ws) -ForegroundColor Red;exit}

try{
  $app=New-Object -ComObject SigmaNEST.SNApp -ErrorAction Stop
  Write-Host ('SNApp.Version = '+$app.Version)
  Write-Host ('Before load: PartsList.Count = '+$app.PartsList.Count)
  $app.LoadWorkSpaceFile($ws)
  Write-Host ('After load: PartsList.Count = '+$app.PartsList.Count) -ForegroundColor Green
  Write-Host ('TasksList.Count = '+$app.TasksList.Count)
  Write-Host ''

  if($app.PartsList.Count -gt 0){
    $part=$app.PartsList.Items(0)
    Write-Host '[ISNPartObj - first part]' -ForegroundColor Yellow
    $part | Get-Member -MemberType Methods,Properties | ForEach-Object {Write-Host ('  '+$_.Name+' :: '+$_.Definition)}
    Write-Host ''
    foreach($p in @('Name','PartName','Number','PartNumber','Material','Thickness','Quantity','Qty','Description','FileName','Path','SourceFile')){
      try{Write-Host ('  VALUE '+$p+' = '+$part.$p)}catch{}
    }
  }else{Write-Host 'No parts found in the loaded workspace.' -ForegroundColor Yellow}

  if($app.TasksList.Count -gt 0){
    $task=$app.TasksList.Items(0)
    Write-Host ''
    Write-Host '[ISNTaskObj - first task]' -ForegroundColor Yellow
    $task | Get-Member -MemberType Methods,Properties | ForEach-Object {Write-Host ('  '+$_.Name+' :: '+$_.Definition)}
  }

  Write-Host ''
  Write-Host '[SNPaths useful values]' -ForegroundColor Yellow
  try{$paths=New-Object -ComObject SigmaNEST.SNPaths;foreach($id in 0..40){try{$v=$paths.GetPath($id);if($v){Write-Host ('  PathID '+$id+' = '+$v)}}catch{}}}catch{Write-Host '  Could not enumerate PathID values.'}

}catch{Write-Host ('ERROR: '+$_.Exception.Message) -ForegroundColor Red}
Write-Host ''
Write-Host '=== END ===' -ForegroundColor Cyan
Write-Host 'Copy all output and send it back.'