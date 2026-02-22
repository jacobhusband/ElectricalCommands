[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [string]$SolutionPath = (Join-Path $PSScriptRoot "ElectricalCommands.sln"),
    [string]$SourceRoot = "$env:APPDATA\Autodesk\ApplicationPlugins",
    [string]$DotnetPath = ""
)

$ErrorActionPreference = "Stop"

function Resolve-DotnetPath {
    param([string]$PathOverride)

    if ($PathOverride) {
        if (-not (Test-Path -LiteralPath $PathOverride)) {
            throw "The provided dotnet path does not exist: $PathOverride"
        }
        return (Resolve-Path -LiteralPath $PathOverride).Path
    }

    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($dotnetCommand) {
        return $dotnetCommand.Source
    }

    throw "Unable to find 'dotnet' in PATH. Install .NET SDK or pass -DotnetPath <path-to-dotnet.exe>."
}

if (-not ([System.IO.Path]::IsPathRooted($SolutionPath))) {
    $SolutionPath = Join-Path $PSScriptRoot $SolutionPath
}
$SolutionPath = [System.IO.Path]::GetFullPath($SolutionPath)
if (-not (Test-Path -LiteralPath $SolutionPath)) {
    throw "Solution file not found: $SolutionPath"
}

if (-not ([System.IO.Path]::IsPathRooted($SourceRoot))) {
    $SourceRoot = Join-Path $PSScriptRoot $SourceRoot
}
$SourceRoot = [System.IO.Path]::GetFullPath($SourceRoot)
New-Item -ItemType Directory -Force -Path $SourceRoot | Out-Null

$dotnetExe = Resolve-DotnetPath -PathOverride $DotnetPath

Write-Host "=== Build Solution ==="
Write-Host "dotnet      : $dotnetExe"
Write-Host "Solution    : $SolutionPath"
Write-Host "Configuration: $Configuration"
Write-Host "BundleRoot  : $SourceRoot"
Write-Host ""

& $dotnetExe build $SolutionPath -c $Configuration "-p:AutoCADBundleRoot=$SourceRoot"
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build failed with exit code $LASTEXITCODE"
}

# Build the AutoLISP bundle so build-bundles.ps1 can zip it with the others.
$autoLispScript = Join-Path $PSScriptRoot "AutoCADCommands\AutoLispCommands\build-autolisp-bundle.ps1"
if (Test-Path -LiteralPath $autoLispScript) {
    & $autoLispScript -TargetRoot $SourceRoot
    if ($LASTEXITCODE -ne 0) {
        throw "AutoLISP bundle build failed with exit code $LASTEXITCODE"
    }
}
