[CmdletBinding()]
param(
  [string]$SourceDir = $PSScriptRoot,
  [string]$TargetRoot = "$env:APPDATA\Autodesk\ApplicationPlugins",
  [string]$BundleName = "ElectricalCommands.AutoLispCommands",
  [string]$Version = ""
)

$ErrorActionPreference = "Stop"

function Get-VersionFromProps {
  param([string]$VersionOverride)

  if ($VersionOverride) { return $VersionOverride }

  $propsPath = Join-Path $PSScriptRoot "..\\..\\version.props"
  if (Test-Path $propsPath) {
    try {
      $xml = [xml](Get-Content $propsPath)
      $verNode = $xml.Project.PropertyGroup.Version
      if ($verNode) { return $verNode }
    }
    catch {
      Write-Warning "Failed to read Version from '$propsPath': $($_.Exception.Message)"
    }
  }

  return "0.0.0"
}

if (-not $SourceDir) { $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$Version = Get-VersionFromProps -VersionOverride $Version

$bundleDir = Join-Path $TargetRoot "$BundleName.bundle"
$contentsDir = Join-Path $bundleDir "Contents"

Write-Host "=== Building AutoLISP bundle (XML Method) ==="
Write-Host "SourceDir : $SourceDir"
Write-Host "BundleDir : $bundleDir"
Write-Host "Version   : $Version"
Write-Host ""

# 1. Clean and Recreate Directories
if (Test-Path $bundleDir) { Remove-Item $bundleDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $contentsDir | Out-Null

# 2. Copy LSP files
# We copy your main loader and the individual utilities
Get-ChildItem -Path $SourceDir -Filter *.lsp | Copy-Item -Destination $contentsDir -Force

# 3. Generate PackageContents.xml
# CHANGES MADE:
# - Removed the standalone <SupportPaths> node (invalid schema location).
# - Added a <ComponentEntry> that points directly to your main loader LSP.
# - Set PerDocument="True" so it mimics acaddoc.lsp behavior (loads on every drawing open).
# - Kept RuntimeRequirements SupportPath="./Contents" so your loader can find the other files via (findfile).

$packageContents = @"
<?xml version="1.0" encoding="utf-8"?>
<ApplicationPackage
  SchemaVersion="1.0"
  AutodeskProduct="AutoCAD"
  ProductType="Application"
  Name="$BundleName"
  Description="AutoLISP utilities for layer toggles, revision clouds, attributes, and text helpers."
  AppVersion="$Version"
  Author="ACIES Engineering"
  ProductCode="{8704562b-e29a-4b5e-8a37-e6db4d303cc0}"
  UpgradeCode="{b1d4f24c-3075-4c76-9e16-b407098b150c}">
  
  <CompanyDetails
    Name="ACIES Engineering"
    Url="www.acies.engineering"
    Email="jacob.h@acies.engineering" />

  <Components>
    <RuntimeRequirements OS="Win64" Platform="AutoCAD*" SupportPath="./Contents" />
    
    <!-- This tells AutoCAD to load your main loader file every time a doc opens -->
    <ComponentEntry
      AppName="ElectricalLoader"
      ModuleName="./Contents/ElectricalCommands.AutoLispCommands.lsp"
      PerDocument="True" />
  </Components>

</ApplicationPackage>
"@

Set-Content -Path (Join-Path $bundleDir "PackageContents.xml") -Value $packageContents -Encoding UTF8
Set-Content -Path (Join-Path $bundleDir "version.txt") -Value ("v{0}" -f $Version) -Encoding UTF8

Write-Host "Bundle created successfully."
