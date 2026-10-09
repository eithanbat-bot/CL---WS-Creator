$ErrorActionPreference='Stop'
$node=Get-Command node.exe -ErrorAction Stop
$test=Join-Path $PSScriptRoot 'test-cl-parser.js'
& $node.Source $test
if($LASTEXITCODE -ne 0){throw "CL parser test failed with exit code $LASTEXITCODE"}
