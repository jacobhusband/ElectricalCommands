param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$InstallerArguments
)

$ErrorActionPreference = "Stop"

try {
    $projectRoot = Split-Path $PSScriptRoot -Parent
    $installerPath = Join-Path $projectRoot "dist\setup\acies-scheduler-setup.exe"

    if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
        throw "Installer not found at $installerPath. Run build first, then run install."
    }

    Write-Host "Launching installer: $installerPath" -ForegroundColor Cyan

    $startProcessArgs = @{
        FilePath = $installerPath
        WorkingDirectory = Split-Path $installerPath -Parent
        Wait = $true
        PassThru = $true
    }

    if ($InstallerArguments.Count -gt 0) {
        $startProcessArgs.ArgumentList = $InstallerArguments
        Write-Host "Forwarding installer arguments: $($InstallerArguments -join ' ')" -ForegroundColor Gray
    }

    $installerProcess = Start-Process @startProcessArgs
    if ($installerProcess.ExitCode -ne 0) {
        throw "Installer exited with code $($installerProcess.ExitCode)."
    }

    Write-Host "Installer completed." -ForegroundColor Green
} catch {
    Write-Error $_
    exit 1
}
