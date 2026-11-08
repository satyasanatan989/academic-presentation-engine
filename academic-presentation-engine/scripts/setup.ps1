# Academic Presentation Engine - Setup Script (setup.ps1)
$ErrorActionPreference = "Stop"

Write-Host "`n=======================================================" -ForegroundColor Cyan
Write-Host " Setting up Academic Presentation Engine Environment" -ForegroundColor Cyan
Write-Host "=======================================================`n" -ForegroundColor Cyan

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $projectRoot

# 1. Create .venv with uv if missing
$venvPath = Join-Path $projectRoot ".venv"
if (-not (Test-Path $venvPath)) {
    Write-Host "[1/4] Creating virtual environment (.venv) using uv..." -ForegroundColor Yellow
    uv venv $venvPath --python 3.13
} else {
    Write-Host "[1/4] Virtual environment already exists at $venvPath" -ForegroundColor Green
}

# 2. Install dependencies
Write-Host "[2/4] Installing / synchronizing pinned dependencies..." -ForegroundColor Yellow
$pyExe = Join-Path $venvPath "Scripts\python.exe"
uv pip install --python $pyExe -r requirements.txt

# 3. Export JSON contracts from Pydantic models
Write-Host "[3/4] Exporting and validating JSON contracts..." -ForegroundColor Yellow
& $pyExe "schemas\export_schemas.py"

# 4. Initialize storage and cache directories
Write-Host "[4/4] Verifying directory hierarchy..." -ForegroundColor Yellow
$dirs = @("storage\cache", "storage\logs", "storage\projects", "outputs")
foreach ($d in $dirs) {
    $fullPath = Join-Path $projectRoot $d
    if (-not (Test-Path $fullPath)) {
        New-Item -ItemType Directory -Path $fullPath -Force | Out-Null
    }
}

Write-Host "`n Setup completed successfully!" -ForegroundColor Green
Write-Host " You can run diagnostics with: .\scripts\doctor.ps1"
Write-Host " You can run the pipeline with:    .\scripts\run.ps1 -Topic 'Your Topic'`n"
