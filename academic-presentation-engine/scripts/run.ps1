# Academic Presentation Engine - Execution Script (run.ps1)
param (
    [Parameter(Mandatory=$true)]
    [string]$Topic,

    [string]$Discipline = "Food Science & Bioprocess Engineering",
    [int]$Slides = 13,
    [int]$Sources = 25,
    [string]$CitationStyle = "APA",
    [string]$Audience = "Academic & Scientific Peers",
    [string]$ProjectId = ""
)

$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $projectRoot

$pyExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $pyExe)) {
    Write-Host "[ERROR] Virtual environment not found at $pyExe. Run .\scripts\setup.ps1 first." -ForegroundColor Red
    exit 1
}

$cmdArgs = @(
    "main.py",
    "--topic", $Topic,
    "--discipline", $Discipline,
    "--slides", $Slides,
    "--sources", $Sources,
    "--citation-style", $CitationStyle,
    "--audience", $Audience
)

if ($ProjectId -ne "") {
    $cmdArgs += @("--project-id", $ProjectId)
}

& $pyExe $cmdArgs
if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[ERROR] Pipeline run failed with exit code $LASTEXITCODE." -ForegroundColor Red
    exit $LASTEXITCODE
}
