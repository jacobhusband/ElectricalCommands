Param(
    # Location of built .bundle folders. Default aligns with project .csproj settings.
    [string]$SourceRoot = "$env:APPDATA\Autodesk\ApplicationPlugins",

    # Where to write zips (relative or absolute). Default: repo AutoCADCommands/dist.
    [string]$OutputRoot = "AutoCADCommands\\dist",

    # Optional explicit version override. If empty, script reads from version.props.
    [string]$Version = "",

    # Build settings forwarded to build.ps1.
    [string]$Configuration = "Release",
    [string]$SolutionPath = (Join-Path $PSScriptRoot "ElectricalCommands.sln"),
    [string]$DotnetPath = ""
)

$ErrorActionPreference = "Stop"

$buildScript = Join-Path $PSScriptRoot "build.ps1"
$bundleScript = Join-Path $PSScriptRoot "build-bundles.ps1"

if (-not (Test-Path $buildScript)) {
    throw "Missing build script at $buildScript"
}
if (-not (Test-Path $bundleScript)) {
    throw "Missing bundle packaging script at $bundleScript"
}

Write-Host "=== Build + Package All Bundles ==="
Write-Host "Build script : $buildScript"
Write-Host "SolutionPath : $SolutionPath"
Write-Host "SourceRoot   : $SourceRoot"
Write-Host "OutputRoot   : $OutputRoot"
Write-Host ""

# Build .NET bundles + AutoLISP bundle into SourceRoot.
& $buildScript `
    -Configuration $Configuration `
    -SolutionPath $SolutionPath `
    -SourceRoot $SourceRoot `
    -DotnetPath $DotnetPath
if ($LASTEXITCODE -ne 0) { exit 1 }

# Zip all bundles into OutputRoot.
& $bundleScript -SourceRoot $SourceRoot -OutputRoot $OutputRoot -Version $Version
