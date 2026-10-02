$ErrorActionPreference='Stop'
$scriptPath=Join-Path $PWD 'bridge/dxf-indexer.ps1'
if(-not(Test-Path -LiteralPath $scriptPath)){throw 'DXF indexer script not found.'}

$root=Join-Path ([IO.Path]::GetTempPath()) ('clwsc-dxf-test '+[Guid]::NewGuid().ToString('N'))
$runtime=Join-Path $root 'runtime'
$indexFile=Join-Path $runtime 'dxf-index.json'
$statusFile=Join-Path $runtime 'dxf-index-status.json'
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

  $requestFile=Join-Path $runtime 'dxf-index-request.json'
  @{Root=$root;IndexFile=$indexFile;StatusFile=$statusFile}|ConvertTo-Json|Set-Content -LiteralPath $requestFile -Encoding UTF8

  & (Get-Command pwsh -ErrorAction Stop).Source -NoProfile -ExecutionPolicy Bypass -File $scriptPath -RequestFile $requestFile
  if($LASTEXITCODE -ne 0){throw "DXF indexer exited with code $LASTEXITCODE"}
  if(Test-Path -LiteralPath $requestFile){throw 'DXF indexer did not consume its request file.'}

  $status=Get-Content -LiteralPath $statusFile -Raw -Encoding UTF8|ConvertFrom-Json
  if([string]$status.state -ne 'COMPLETE'){throw "Unexpected final status: $($status.state) / $($status.message)"}
  if([string]$status.indexerVersion -ne '2.1.0'){throw "Unexpected indexer version: $($status.indexerVersion)"}
  if([int]$status.filesFound -ne 3){throw "Expected 3 DXF files, found $($status.filesFound)"}
  if(-not(Test-Path -LiteralPath $indexFile)){throw 'DXF index manifest was not created.'}
  $shardDir=Join-Path $runtime 'dxf-index'
  if(-not(Test-Path -LiteralPath $shardDir)){throw 'DXF shard directory was not created.'}

  $manifest=Get-Content -LiteralPath $indexFile -Raw -Encoding UTF8|ConvertFrom-Json
  if([int]$manifest.filesFound -ne 3){throw 'Manifest file count mismatch.'}
  if([int]$manifest.directoriesVisited -lt 3){throw 'Expected recursive directory walk did not occur.'}

  $lines=@()
  Get-ChildItem -LiteralPath $shardDir -Filter '*.tsv' -File|ForEach-Object{
    $lines += @(Get-Content -LiteralPath $_.FullName -Encoding UTF8)
  }
  $allText=$lines -join [Environment]::NewLine
  foreach($f in $files){
    if(-not $allText.Contains($f)){throw "Indexed path missing: $f"}
  }

  Write-Host 'DXF indexer end-to-end test: OK' -ForegroundColor Green
}finally{
  Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
