$ErrorActionPreference='Stop'
$scriptPath=Join-Path $PWD 'bridge/dxf-indexer.ps1'
if(-not(Test-Path -LiteralPath $scriptPath)){throw 'DXF indexer script not found.'}

$root=Join-Path ([IO.Path]::GetTempPath()) ('clwsc-dxf-test '+[Guid]::NewGuid().ToString('N'))
$runtime=Join-Path $root 'runtime'
New-Item -ItemType Directory -Path $runtime -Force|Out-Null
$indexFile=Join-Path $runtime 'dxf-index.json'
$statusFile=Join-Path $runtime 'dxf-index-status.json'

function Invoke-Indexer([string]$mode){
  $requestFile=Join-Path $runtime ('dxf-index-request-'+$mode.ToLowerInvariant()+'.json')
  @{Role='CONTROLLER';Mode=$mode;WorkerCount=2;Root=$root;IndexFile=$indexFile;StatusFile=$statusFile}|ConvertTo-Json|Set-Content -LiteralPath $requestFile -Encoding UTF8
  & (Get-Command pwsh -ErrorAction Stop).Source -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Role CONTROLLER -RequestFile $requestFile
  if($LASTEXITCODE -ne 0){throw "DXF indexer exited with code $LASTEXITCODE in $mode mode"}
  return (Get-Content -LiteralPath $statusFile -Raw -Encoding UTF8|ConvertFrom-Json)
}

try{
  $dirA=Join-Path $root 'A Folder'
  $dirNested=Join-Path $dirA 'Nested One'
  $dirB=Join-Path $root 'B Folder'
  New-Item -ItemType Directory -Path $dirNested -Force|Out-Null
  New-Item -ItemType Directory -Path $dirB -Force|Out-Null

  $files=@(
    (Join-Path $dirA 'one.DXF'),
    (Join-Path $dirNested 'two.dxf'),
    (Join-Path $dirB 'three.DxF')
  )
  foreach($f in $files){Set-Content -LiteralPath $f -Value '0' -Encoding ASCII}
  Set-Content -LiteralPath (Join-Path $dirB 'ignore.txt') -Value '0' -Encoding ASCII

  $status=Invoke-Indexer -mode 'FULL'
  if([string]$status.state -ne 'COMPLETE'){throw "Unexpected FULL status: $($status.state) / $($status.message)"}
  if([string]$status.indexerVersion -ne '3.0.0'){throw "Unexpected indexer version: $($status.indexerVersion)"}
  if([string]$status.mode -ne 'FULL'){throw "Expected FULL mode, got $($status.mode)"}
  if([int]$status.filesFound -ne 3){throw "Expected 3 DXF files, found $($status.filesFound)"}
  if([int]$status.workers -ne 2){throw "Expected 2 worker(s), found $($status.workers)"}
  if([int]$status.workersCompleted -ne 2){throw "Expected all workers to complete."}
  if(-not(Test-Path -LiteralPath $indexFile)){throw 'DXF index manifest was not created.'}
  $shardDir=Join-Path $runtime 'dxf-index'
  if(-not(Test-Path -LiteralPath $shardDir)){throw 'DXF shard directory was not created.'}

  $manifest=Get-Content -LiteralPath $indexFile -Raw -Encoding UTF8|ConvertFrom-Json
  if([string]$manifest.schema -ne 'cl-ws-creator/dxf-index/3'){throw "Unexpected manifest schema: $($manifest.schema)"}
  if([int]$manifest.filesFound -ne 3){throw 'Manifest file count mismatch.'}
  if([int]$manifest.workers -ne 2){throw 'Manifest worker count mismatch.'}

  $allText=''
  Get-ChildItem -LiteralPath $shardDir -Filter '*.tsv' -File|ForEach-Object{$allText += (Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8)}
  foreach($f in $files){
    if(-not $allText.Contains($f)){throw "Indexed path missing after FULL scan: $f"}
  }

  $newFile=Join-Path $dirB 'four.DXF'
  Set-Content -LiteralPath $newFile -Value '1' -Encoding ASCII
  $status=Invoke-Indexer -mode 'REFRESH'
  if([string]$status.state -ne 'COMPLETE'){throw "Unexpected REFRESH status: $($status.state) / $($status.message)"}
  if([string]$status.mode -ne 'REFRESH'){throw "Expected REFRESH mode, got $($status.mode)"}
  if([int]$status.filesFound -ne 4){throw "Expected 4 DXF files after refresh, found $($status.filesFound)"}

  $manifest=Get-Content -LiteralPath $indexFile -Raw -Encoding UTF8|ConvertFrom-Json
  if([string]$manifest.mode -ne 'REFRESH'){throw 'Refresh manifest mode was not persisted.'}
  if([int]$manifest.filesFound -ne 4){throw 'Refresh manifest count mismatch.'}

  Write-Host 'DXF indexer parallel full + refresh test: OK' -ForegroundColor Green
}finally{
  Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
