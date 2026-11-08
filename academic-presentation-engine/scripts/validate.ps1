# Academic Presentation Engine - Validation Script (validate.ps1)
$ErrorActionPreference = "Continue"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $projectRoot

Write-Host "`n=======================================================" -ForegroundColor Cyan
Write-Host " Executing Complete Quality Assurance & Test Validation" -ForegroundColor Cyan
Write-Host "=======================================================`n" -ForegroundColor Cyan

$pytestExe = Join-Path $projectRoot ".venv\Scripts\pytest.exe"
if (-not (Test-Path $pytestExe)) {
    Write-Host "[ERROR] pytest not found. Run .\scripts\setup.ps1 first." -ForegroundColor Red
    exit 1
}

& $pytestExe -v tests
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Host "`n [PASS] All test suites and quality gates passed successfully." -ForegroundColor Green
} else {
    Write-Host "`n [FAIL] Some validations failed with exit code $exitCode." -ForegroundColor Red
}
exit $exitCode
