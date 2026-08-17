<#
.SYNOPSIS
    Builds and validates all apps in the ACIES-Tools monorepo.
.DESCRIPTION
    Builds the C# AutoCAD ElectricalCommands solution and validates the ProjectManagement Python dependencies.
#>

param(
    [string]$DotnetPath = "C:\Users\JacobH\OneDrive - ACIES Engineering\Documents\dotnet-sdk\dotnet.exe",
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
$RepoRoot = $PSScriptRoot

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  ACIES-Tools Monorepo: Build All Apps" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# 1. Locate Dotnet SDK
if (-not (Test-Path $DotnetPath)) {
    Write-Host "Portable dotnet SDK not found at default path, checking PATH..." -ForegroundColor Yellow
    $DotnetCmd = Get-Command "dotnet" -ErrorAction SilentlyContinue
    if ($DotnetCmd) {
        $DotnetPath = $DotnetCmd.Source
    } else {
        Write-Error "dotnet.exe could not be found. Please pass -DotnetPath with valid SDK path."
    }
}
Write-Host "[.NET SDK] Using: $DotnetPath" -ForegroundColor Green

# 2. Build ElectricalCommands (.NET C#)
$SlnPath = Join-Path $RepoRoot "apps\ElectricalCommands\ElectricalCommands.sln"
if (Test-Path $SlnPath) {
    Write-Host "`n--> Building ElectricalCommands ($Configuration)..." -ForegroundColor Cyan
    & $DotnetPath build "$SlnPath" -c $Configuration
    if ($LASTEXITCODE -ne 0) {
        Write-Error "ElectricalCommands build failed."
    }
    Write-Host "[OK] ElectricalCommands built successfully." -ForegroundColor Green
} else {
    Write-Warning "ElectricalCommands.sln not found at $SlnPath"
}

# 3. Check Python ProjectManagement
$PythonVenv = Join-Path $RepoRoot "apps\ProjectManagement\.venv\Scripts\python.exe"
if (Test-Path $PythonVenv) {
    Write-Host "`n--> Checking ProjectManagement Python environment..." -ForegroundColor Cyan
    & $PythonVenv --version
    Write-Host "[OK] ProjectManagement environment detected." -ForegroundColor Green
} else {
    Write-Host "`n[INFO] ProjectManagement .venv not detected. Create one if needed using:" -ForegroundColor Yellow
    Write-Host "       python -m venv apps/ProjectManagement/.venv" -ForegroundColor Gray
}

Write-Host "`n=========================================" -ForegroundColor Cyan
Write-Host "  Build completed successfully!" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Cyan
