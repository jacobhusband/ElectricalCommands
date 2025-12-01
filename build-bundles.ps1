Param(
    # Where to read bundles from:
    # - By default, from the installed ApplicationPlugins directory.
    # - Override with -SourceRoot "AutoCADCommands" if you keep .bundle folders in the repo.
    [string]$SourceRoot = "$env:APPDATA\Autodesk\ApplicationPlugins",

    # Where to write zips (relative or absolute). Default: repo AutoCADCommands/dist.
    [string]$OutputRoot = "AutoCADCommands\dist",

    # Optional explicit version override. If empty, script reads from a csproj.
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

Write-Host "=== AutoCAD Bundle Packaging ==="
Write-Host "SourceRoot : $SourceRoot"
Write-Host "OutputRoot:  $OutputRoot"
Write-Host ""

if (-not (Test-Path $SourceRoot)) {
    Write-Host "SourceRoot '$SourceRoot' does not exist. Creating it..."
    New-Item -ItemType Directory -Force -Path $SourceRoot | Out-Null
}

# Normalize output directory relative to current working directory if needed
if (-not ([System.IO.Path]::IsPathRooted($OutputRoot))) {
    $OutputRoot = Join-Path (Get-Location) $OutputRoot
}
New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null

function New-Zip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$ZipPath
    )

    if (-not (Test-Path $SourcePath)) {
        throw "SourcePath '$SourcePath' does not exist."
    }

    if (Test-Path $ZipPath) {
        Remove-Item $ZipPath -Force
    }

    Add-Type -AssemblyName "System.IO.Compression.FileSystem" -ErrorAction Stop
    [System.IO.Compression.ZipFile]::CreateFromDirectory($SourcePath, $ZipPath)
}

function New-AutoLispBundle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,
        [Parameter(Mandatory = $true)]
        [string]$TargetRoot,
        [string]$BundleName = "ElectricalCommands.AutoLispCommands",
        [string]$Version = "0.0.0"
    )

    if (-not (Test-Path $SourceDir)) {
        Write-Warning "AutoLISP source directory '$SourceDir' not found. Skipping AutoLISP bundle build."
        return
    }

    if (-not (Test-Path $TargetRoot)) {
        New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null
    }

    $bundleDir = Join-Path $TargetRoot "$BundleName.bundle"
    $contentsDir = Join-Path $bundleDir "Contents"

    Write-Host "=== Building AutoLISP bundle ==="
    Write-Host "SourceDir : $SourceDir"
    Write-Host "BundleDir : $bundleDir"
    Write-Host "Version   : $Version"
    Write-Host ""

    if (Test-Path $bundleDir) { Remove-Item $bundleDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $contentsDir | Out-Null

    Get-ChildItem -Path $SourceDir -Filter *.lsp -File | Copy-Item -Destination $contentsDir -Force

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
    
    <ComponentEntry
      AppName="ElectricalLoader"
      ModuleName="./Contents/ElectricalCommands.AutoLispCommands.lsp"
      PerDocument="True" />
  </Components>

</ApplicationPackage>
"@

    Set-Content -Path (Join-Path $bundleDir "PackageContents.xml") -Value $packageContents -Encoding UTF8
    Set-Content -Path (Join-Path $bundleDir "version.txt") -Value ("v{0}" -f $Version) -Encoding UTF8

    Write-Host "[done] AutoLISP bundle created at $bundleDir"
    Write-Host ""
}

# Map: bundle folder -> zip name prefix
$bundles = @(
    @{ Folder = "ElectricalCommands.AutoLispCommands.bundle"; Prefix = "ElectricalCommands.AutoLispCommands" },
    @{ Folder = "ElectricalCommands.CleanCADCommands.bundle"; Prefix = "ElectricalCommands.CleanCADCommands" },
    @{ Folder = "ElectricalCommands.GeneralCommands.bundle"; Prefix = "ElectricalCommands.GeneralCommands" },
    @{ Folder = "ElectricalCommands.GetAttributesCommands.bundle"; Prefix = "ElectricalCommands.GetAttributesCommands" },
    @{ Folder = "ElectricalCommands.PlotCommands.bundle"; Prefix = "ElectricalCommands.PlotCommands" },
    @{ Folder = "ElectricalCommands.T24Commands.bundle"; Prefix = "ElectricalCommands.T24Commands" },
    @{ Folder = "ElectricalCommands.TextCommands.bundle"; Prefix = "ElectricalCommands.TextCommands" }
)

# Determine version if not explicitly provided
if (-not $Version) {
    $Version = "0.0.0"
    # Look for a csproj in the repo structure above/around OutputRoot
    $repoRoot = (Get-Location)
    $csproj = Get-ChildItem -Path $repoRoot -Recurse -Filter "*.csproj" | Select-Object -First 1
    if ($csproj) {
        try {
            $xml = [xml](Get-Content $csproj.FullName)
            $verNode = $xml.Project.PropertyGroup.Version
            if ($verNode) {
                $Version = $verNode
            }
        }
        catch {
            Write-Warning "Failed to read Version from '$($csproj.FullName)': $($_.Exception.Message)"
        }
    }
}

Write-Host "Using version: $Version"
Write-Host ""

# Build AutoLISP bundle (it doesn't have its own project build step)
$autoLispSourceDir = Join-Path $PSScriptRoot "AutoCADCommands\AutoLispCommands"
New-AutoLispBundle -SourceDir $autoLispSourceDir -TargetRoot $SourceRoot -Version $Version

# Collect metadata from description files
$metadata = @{}
$descriptionFiles = Get-ChildItem -Path "AutoCADCommands" -Recurse -Filter "*_descriptions.json"
foreach ($file in $descriptionFiles) {
    try {
        $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $bundleName = $file.Directory.Name
        $prefix = "ElectricalCommands.$bundleName"
        $metadata[$prefix] = @{
            video    = $json.video
            commands = $json.commands
        }
    }
    catch {
        Write-Warning "Failed to parse $($file.FullName): $($_.Exception.Message)"
    }
}

foreach ($b in $bundles) {
    $bundlePath = Join-Path $SourceRoot $b.Folder

    if (-not (Test-Path $bundlePath)) {
        Write-Host ("[skip] {0}: bundle not found at {1}" -f $b.Folder, $bundlePath)
        continue
    }

    $zipName = "{0}-v{1}.zip" -f $b.Prefix, $Version
    $zipPath = Join-Path $OutputRoot $zipName

    Write-Host ("[pack] {0} -> {1}" -f $bundlePath, $zipPath)
    New-Zip -SourcePath $bundlePath -ZipPath $zipPath
}

# Generate release_meta.json
$assets = @()
foreach ($b in $bundles) {
    $zipName = "{0}-v{1}.zip" -f $b.Prefix, $Version
    $commands = @{}
    $video = $null
    if ($metadata.ContainsKey($b.Prefix)) {
        $commands = $metadata[$b.Prefix].commands
        $video = $metadata[$b.Prefix].video
    }
    $assets += @{
        filename  = $zipName
        video_url = $video
        commands  = $commands
    }
}

$releaseMeta = @{ assets = $assets } | ConvertTo-Json -Depth 10
$metaPath = Join-Path $OutputRoot "release_meta.json"
$releaseMeta | Out-File $metaPath -Encoding UTF8

Write-Host ""
Write-Host "Done. Zips and metadata are in: $OutputRoot"
Write-Host "Upload these as GitHub Release assets."
