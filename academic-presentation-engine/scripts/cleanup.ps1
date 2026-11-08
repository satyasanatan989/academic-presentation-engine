# Academic Presentation Engine - Safe Cleanup Script (cleanup.ps1)
$ErrorActionPreference = "Continue"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $projectRoot

Write-Host "`n=======================================================" -ForegroundColor Cyan
Write-Host " Safe Cleanup: Temporary Files, Caches & Bytecode" -ForegroundColor Cyan
Write-Host "=======================================================`n" -ForegroundColor Cyan

# 1. Clean Python __pycache__ and bytecode
$pycacheDirs = Get-ChildItem -Path $projectRoot -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue
foreach ($dir in $pycacheDirs) {
    if ($dir.FullName -notmatch "\.venv") {
        Write-Host "Removing: $($dir.FullName)" -ForegroundColor DarkGray
        Remove-Item -Path $dir.FullName -Recurse -Force | Out-Null
    }
}

# 2. Clean .pytest_cache
$pytestCache = Join-Path $projectRoot ".pytest_cache"
if (Test-Path $pytestCache) {
    Write-Host "Removing: $pytestCache" -ForegroundColor DarkGray
    Remove-Item -Path $pytestCache -Recurse -Force | Out-Null
}

# 3. Clean storage/cache files (preserve .gitkeep)
$cacheFiles = Get-ChildItem -Path (Join-Path $projectRoot "storage\cache") -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne ".gitkeep" }
foreach ($f in $cacheFiles) {
    Write-Host "Clearing cache: $($f.Name)" -ForegroundColor DarkGray
    Remove-Item -Path $f.FullName -Force | Out-Null
}

Write-Host "`n Cleanup completed safely. Deliverables, configurations, and .venv preserved.`n" -ForegroundColor Green
