# Academic Presentation Engine - System & Dependency Diagnostics (doctor.ps1)
$ErrorActionPreference = "Continue"

Write-Host "`n=======================================================" -ForegroundColor Cyan
Write-Host " Academic Presentation Engine: System Doctor" -ForegroundColor Cyan
Write-Host "=======================================================`n" -ForegroundColor Cyan

$allPassed = $true

# 1. OS & Architecture Check
$os = Get-CimInstance Win32_OperatingSystem
Write-Host "[OK] Operating System: $($os.Caption) (Build $($os.BuildNumber), 64-bit)" -ForegroundColor Green

# 2. RAM Check
$ramGb = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
if ($ramGb -ge 7.0) {
    Write-Host "[OK] Physical RAM: $ramGb GB (Optimal for low-memory pipeline)" -ForegroundColor Green
} else {
    Write-Host "[WARN] Physical RAM: $ramGb GB (Minimum 8GB recommended)" -ForegroundColor Yellow
}

# 3. Python 3.12+ Check
try {
    $pyVer = python --version 2>&1
    if ($pyVer -match "Python 3\.(1[2-9]|[2-9][0-9])") {
        Write-Host "[OK] Python: $pyVer" -ForegroundColor Green
    } else {
        Write-Host "[WARN] Python version may be incompatible: $pyVer (Requires 3.12+)" -ForegroundColor Yellow
        $allPassed = $false
    }
} catch {
    Write-Host "[FAIL] Python not found on PATH." -ForegroundColor Red
    $allPassed = $false
}

# 4. UV Package Manager Check
try {
    $uvVer = uv --version 2>&1
    Write-Host "[OK] UV Package Manager: $uvVer" -ForegroundColor Green
} catch {
    Write-Host "[WARN] UV not found. Standard pip will be used as fallback." -ForegroundColor Yellow
}

# 5. Virtual Environment Check
$venvPy = Join-Path $PSScriptRoot "..\.venv\Scripts\python.exe"
if (Test-Path $venvPy) {
    Write-Host "[OK] Dedicated Virtual Environment: Present ($venvPy)" -ForegroundColor Green
} else {
    Write-Host "[WARN] Virtual environment not found. Run .\scripts\setup.ps1" -ForegroundColor Yellow
    $allPassed = $false
}

# 6. Microsoft PowerPoint 2024 / COM Automation Check
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ver = $ppt.Version
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    Write-Host "[OK] Microsoft PowerPoint COM Automation: Available (Office Version $ver)" -ForegroundColor Green
} catch {
    Write-Host "[FAIL] Microsoft PowerPoint COM not accessible: $_" -ForegroundColor Red
    $allPassed = $false
}

# 7. Disk Space Check
$drive = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -eq "C:\" }
$freeGb = [math]::Round($drive.Free / 1GB, 2)
if ($freeGb -ge 5.0) {
    Write-Host "[OK] Available Disk Space on C:\: $freeGb GB" -ForegroundColor Green
} else {
    Write-Host "[WARN] Low disk space on C:\: $freeGb GB" -ForegroundColor Yellow
}

Write-Host "`n-------------------------------------------------------"
if ($allPassed) {
    Write-Host " Diagnosis: ALL SYSTEMS READY FOR PRODUCTION EXECUTION." -ForegroundColor Green
} else {
    Write-Host " Diagnosis: ISSUES DETECTED. Review notices above." -ForegroundColor Yellow
}
Write-Host "-------------------------------------------------------`n"
