[CmdletBinding()]
param(
  [switch]$RealCom,
  [switch]$HttpE2E,
  [string]$DxfPath='Y:\SOLIDWORKS\UNIVERSAL COMPONENTS\PC-2A.DXF',
  [string]$PrsPath='S:\SNDataX1\PARTS\HSWU2 100.PRS',
  [int]$TestQty=4,
  [string]$TestMaterial='MS',
  [double]$TestThicknessMm=4.0,
  [int]$BatchMultiplier=4
)
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot

$required=@(
  @{Name='Static release gate';Path=(Join-Path $PSScriptRoot 'test-sigmanest-static.ps1')},
  @{Name='Deterministic COM mock';Path=(Join-Path $PSScriptRoot 'test-sigmanest-com-mock.ps1')}
)
foreach($test in $required){
  if(-not(Test-Path -LiteralPath $test.Path)){throw "Required test is missing: $($test.Path)"}
  Write-Host ("=== "+$test.Name+" ===") -ForegroundColor Cyan
  & $test.Path
}

if($RealCom){
  $testPath=Join-Path $PSScriptRoot 'test-sigmanest-adapter-real.ps1'
  if(-not(Test-Path -LiteralPath $testPath)){throw "Real COM test is missing: $testPath"}
  Write-Host '=== Real SigmaNEST COM integration ===' -ForegroundColor Cyan
  & $testPath -DxfPath $DxfPath -TestQty $TestQty -TestMaterial $TestMaterial -TestThicknessMm $TestThicknessMm -BatchMultiplier $BatchMultiplier

  $mixedPath=Join-Path $PSScriptRoot 'test-sigmanest-mixed-real.ps1'
  if(-not(Test-Path -LiteralPath $mixedPath)){throw "Mixed DXF/PRS integration test is missing: $mixedPath"}
  Write-Host '=== Mixed DXF + PRS SigmaNEST integration ===' -ForegroundColor Cyan
  & $mixedPath -DxfPath $DxfPath -PrsPath $PrsPath
}

if($HttpE2E){
  $testPath=Join-Path $PSScriptRoot 'test-sigmanest-http-e2e.ps1'
  if(-not(Test-Path -LiteralPath $testPath)){throw "HTTP end-to-end test is missing: $testPath"}
  Write-Host '=== Live bridge HTTP-to-SigmaNEST end-to-end ===' -ForegroundColor Cyan
  & $testPath
}

Write-Host 'CL - WS Creator release gate: ALL REQUESTED TESTS PASSED' -ForegroundColor Green
