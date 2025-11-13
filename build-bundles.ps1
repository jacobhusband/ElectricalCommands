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
    throw "SourceRoot '$SourceRoot' does not exist. Adjust -SourceRoot."
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

# Map: bundle folder -> zip name prefix
$bundles = @(
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