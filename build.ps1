$ErrorActionPreference = "Stop"

& "C:\Users\JacobH\OneDrive - ACIES Engineering\Documents\dotnet-sdk\dotnet.exe" build "c:\Users\JacobH\Documents\dev\ElectricalCommands\ElectricalCommands.sln" -c Release

# Build the AutoLISP bundle so build-bundles.ps1 can zip it with the others.
$autoLispScript = Join-Path $PSScriptRoot "AutoCADCommands\AutoLispCommands\build-autolisp-bundle.ps1"
if (Test-Path $autoLispScript) {
    & $autoLispScript
}
