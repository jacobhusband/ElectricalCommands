Param(
    [string]$Notes = "Automated release",
    [string]$Version = "",
    [switch]$Build
)

$ErrorActionPreference = "Stop"

# --- Config ---
$repo = "jacobhusband/ElectricalCommands"
$distRoot = Join-Path $PSScriptRoot "AutoCADCommands\\dist"
$versionPropsPath = Join-Path $PSScriptRoot "version.props"

function Import-DotEnv {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    Get-Content $Path | ForEach-Object {
        if ($_ -match '^\s*#') { return }
        if ($_ -match '^\s*$') { return }
        $parts = $_ -split '=', 2
        if ($parts.Count -eq 2) {
            $name = $parts[0].Trim()
            $value = $parts[1].Trim('"').Trim()
            if ($name) { Set-Item -Path "Env:$name" -Value $value }
        }
    }
}

function Get-VersionFromProps {
    param([string]$Override)

    if ($Override) { return $Override }
    if (-not (Test-Path $versionPropsPath)) {
        throw "version.props not found at $versionPropsPath. Pass -Version or add the file."
    }

    $xml = [xml](Get-Content $versionPropsPath)
    $ver = $xml.Project.PropertyGroup.Version
    if (-not $ver) { throw "Version node missing in $versionPropsPath." }
    return $ver
}

# --- Token ---
# Load from .env if present (convenience)
Import-DotEnv -Path (Join-Path $PSScriptRoot ".env")
$token = $env:GITHUB_TOKEN
if (-not $token) {
    Write-Error "Set environment variable GITHUB_TOKEN to a PAT with repo permissions."
    exit 1
}

# --- Version ---
try {
    $version = Get-VersionFromProps -Override $Version
} catch {
    Write-Error $_
    exit 1
}
$tag = "v$version"

# --- Optional build ---
if ($Build) {
    & (Join-Path $PSScriptRoot "build-bundles.ps1") -Version $version
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

if (-not (Test-Path $distRoot)) {
    Write-Error "dist folder not found at $distRoot. Run build-bundles.ps1 or pass -Build."
    exit 1
}

# Collect assets: all zip files and release_meta.json if present
$assetsToUpload = @()
$zipFiles = Get-ChildItem -Path $distRoot -Filter "*.zip" -File
if (-not $zipFiles) {
    Write-Error "No zip assets found in $distRoot. Build first or pass -Build."
    exit 1
}
$assetsToUpload += $zipFiles
$metaPath = Join-Path $distRoot "release_meta.json"
if (Test-Path $metaPath) {
    $assetsToUpload += Get-Item $metaPath
}

$headers = @{
    "Authorization" = "token $token"
    "Accept"        = "application/vnd.github+json"
    "User-Agent"    = "publish-release-script"
}

# --- Create release if missing ---
$release = $null
try {
    $release = Invoke-RestMethod -Method Get -Headers $headers -Uri "https://api.github.com/repos/$repo/releases/tags/$tag"
} catch { $release = $null }

if (-not $release) {
    $body = @{
        tag_name   = $tag
        name       = $tag
        body       = $Notes
        draft      = $false
        prerelease = $false
    } | ConvertTo-Json
    $release = Invoke-RestMethod -Method Post -Headers $headers -Uri "https://api.github.com/repos/$repo/releases" -Body $body
}

if (-not $release.upload_url) {
    Write-Error "Release upload_url missing. Check token permissions (repo:write) and try again."
    exit 1
}

$uploadUrl = [string]$release.upload_url
$uploadUrl = $uploadUrl.Replace("{?name,label}", "").Trim()

# --- Upload assets (delete existing matching names) ---
foreach ($asset in $assetsToUpload) {
    $assetName = $asset.Name
    $encodedName = [uri]::EscapeDataString($assetName)
    $uploadUri = "{0}?name={1}" -f $uploadUrl, $encodedName

    try {
        $existing = Invoke-RestMethod -Method Get -Headers $headers -Uri "https://api.github.com/repos/$repo/releases/$($release.id)/assets"
        foreach ($a in $existing) {
            if ($a.name -eq $assetName) {
                Invoke-RestMethod -Method Delete -Headers $headers -Uri "https://api.github.com/repos/$repo/releases/assets/$($a.id)" | Out-Null
            }
        }
    } catch { }

    $assetHeaders = $headers.Clone()
    $assetHeaders["Content-Type"] = "application/octet-stream"
    Write-Host "Uploading $assetName to $uploadUri" -ForegroundColor Cyan
    try {
        Invoke-RestMethod -Method Post -Headers $assetHeaders -Uri $uploadUri -InFile $asset.FullName -TimeoutSec 600 | Out-Null
    } catch {
        Write-Error "Asset upload failed for $assetName. uploadUri: $uploadUri`nError: $_"
        exit 1
    }
}

Write-Host "Release $tag published with $($assetsToUpload.Count) asset(s)." -ForegroundColor Green
