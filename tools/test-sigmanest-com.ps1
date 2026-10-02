$ErrorActionPreference = "Continue"

$progIds = @(
  "SigmaNEST.SNAutomation",
  "SigmaNEST.SNApp",
  "SigmaNEST.SNWorkOrder",
  "SigmaNEST.SNWSDoc",
  "SigmaNEST.SNTask",
  "SigmaNEST.SNTaskDoc",
  "SigmaNEST.SNPartDoc",
  "SigmaNEST.SNPartExportImport",
  "SigmaNEST.SNPaths"
)

Write-Host "=== SigmaNEST COM Diagnostic ===" -ForegroundColor Cyan
Write-Host "Keep SigmaNEST open while this runs."
Write-Host ""

foreach ($progId in $progIds) {
    Write-Host ("[" + $progId + "]") -ForegroundColor Yellow
    try {
        $obj = New-Object -ComObject $progId -ErrorAction Stop
        Write-Host "  CREATE: OK" -ForegroundColor Green
        try {
            $members = $obj | Get-Member -MemberType Methods,Properties -ErrorAction Stop |
                Select-Object -ExpandProperty Name -Unique | Sort-Object
            if ($members) {
                Write-Host "  MEMBERS:"
                $members | ForEach-Object { Write-Host ("    " + $_) }
            } else { Write-Host "  MEMBERS: none returned" }
        } catch {
            Write-Host ("  MEMBERS: ERROR " + $_.Exception.Message) -ForegroundColor Red
        }
    } catch {
        Write-Host "  CREATE: FAILED" -ForegroundColor Red
        Write-Host ("  ERROR : " + $_.Exception.Message) -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "=== END ===" -ForegroundColor Cyan
Write-Host "Copy everything above and send it back to ChatGPT."
