$ErrorActionPreference = "Stop"

function Invoke-CheckedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Command
    )

    $global:LASTEXITCODE = 0
    $output = & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }

    return $output
}

function Assert-PathExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not (Test-Path $Path)) {
        throw $Message
    }
}

function Get-NextPatchVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    if ($Version -notmatch '^(?<major>\d+)\.(?<minor>\d+)\.(?<patch>\d+)$') {
        throw "VERSION must use semantic version format X.Y.Z. Found: $Version"
    }

    $major = [int]$Matches["major"]
    $minor = [int]$Matches["minor"]
    $patch = [int]$Matches["patch"] + 1
    return "$major.$minor.$patch"
}

function Write-Utf8NoBomFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Value, $utf8NoBom)
}

function Convert-RemoteUrlToWebUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RemoteUrl
    )

    if ($RemoteUrl -match '^(https?://)(?<host>[^/]+)/(?<repo>.+?)(?:\.git)?$') {
        return "https://$($Matches["host"])/$($Matches["repo"])"
    }

    if ($RemoteUrl -match '^git@(?<host>[^:]+):(?<repo>.+?)(?:\.git)?$') {
        return "https://$($Matches["host"])/$($Matches["repo"])"
    }

    if ($RemoteUrl -match '^ssh://git@(?<host>[^/]+)/(?<repo>.+?)(?:\.git)?$') {
        return "https://$($Matches["host"])/$($Matches["repo"])"
    }

    return ($RemoteUrl -replace '\.git$', '')
}

$projectRoot = Split-Path $PSScriptRoot -Parent
$buildScriptPath = Join-Path $PSScriptRoot "build.ps1"
$versionPath = Join-Path $projectRoot "VERSION"

$originalVersionText = $null
$originalVersion = $null
$nextVersion = $null
$tag = $null
$versionUpdated = $false
$commitCreated = $false
$locationPushed = $false

try {
    Assert-PathExists -Path $buildScriptPath -Message "Build script not found at $buildScriptPath."
    Assert-PathExists -Path $versionPath -Message "VERSION file not found at $versionPath."

    Push-Location $projectRoot
    $locationPushed = $true

    $gitStatus = @(Invoke-CheckedCommand -Description "git status" -Command { git status --porcelain })
    if ($gitStatus.Count -gt 0) {
        throw "Git working tree is not clean. Commit or stash changes before releasing.`n$($gitStatus -join "`n")"
    }

    $currentBranch = ((Invoke-CheckedCommand -Description "git branch lookup" -Command { git branch --show-current }) | Select-Object -First 1).Trim()
    if ($currentBranch -ne "main") {
        throw "Releases must be created from the main branch. Current branch: $currentBranch"
    }

    Write-Host "Fetching origin/main and tags..." -ForegroundColor Cyan
    Invoke-CheckedCommand -Description "git fetch origin" -Command { git fetch origin main --tags }

    $aheadBehind = ((Invoke-CheckedCommand -Description "git rev-list comparison" -Command { git rev-list --left-right --count HEAD...origin/main }) | Select-Object -First 1).Trim()
    if ($aheadBehind -notmatch '^(?<ahead>\d+)\s+(?<behind>\d+)$') {
        throw "Could not parse ahead/behind counts from git rev-list output: $aheadBehind"
    }

    $aheadCount = [int]$Matches["ahead"]
    $behindCount = [int]$Matches["behind"]
    if ($behindCount -gt 0) {
        throw "Local main is behind origin/main by $behindCount commit(s). Pull or rebase before releasing."
    }

    if ($aheadCount -gt 0) {
        Write-Host "Local main is ahead of origin/main by $aheadCount commit(s)." -ForegroundColor Gray
    }

    $originalVersionText = [System.IO.File]::ReadAllText($versionPath)
    $versionLineEnding = if ($originalVersionText.Contains("`r`n")) { "`r`n" } elseif ($originalVersionText.Contains("`n")) { "`n" } else { [System.Environment]::NewLine }
    $originalVersion = $originalVersionText.Trim()
    if (-not $originalVersion) {
        throw "VERSION file is empty."
    }

    $nextVersion = Get-NextPatchVersion -Version $originalVersion
    $tag = "v$nextVersion"

    $existingTag = ((Invoke-CheckedCommand -Description "git tag lookup" -Command { git tag --list $tag }) | Select-Object -First 1)
    if ($existingTag) {
        throw "Tag $tag already exists. Update VERSION or remove the conflicting tag before releasing."
    }

    Write-Utf8NoBomFile -Path $versionPath -Value ($nextVersion + $versionLineEnding)
    $versionUpdated = $true
    Write-Host "Bumped VERSION from $originalVersion to $nextVersion" -ForegroundColor Yellow

    Write-Host "Running local build gate..." -ForegroundColor Cyan
    Invoke-CheckedCommand -Description "Local build gate" -Command { & $buildScriptPath }

    $postBuildStatus = @(Invoke-CheckedCommand -Description "git status after build" -Command { git status --porcelain })
    $unexpectedChanges = @($postBuildStatus | Where-Object { $_ -notmatch '^.. VERSION$' })
    if ($postBuildStatus.Count -ne 1 -or $unexpectedChanges.Count -gt 0) {
        throw "Build modified unexpected tracked files. Release commit would not be clean.`n$($postBuildStatus -join "`n")"
    }

    Invoke-CheckedCommand -Description "git add VERSION" -Command { git add -- VERSION }
    Invoke-CheckedCommand -Description "git commit VERSION bump" -Command { git commit -m "Bump version to $nextVersion" }
    $commitCreated = $true

    Invoke-CheckedCommand -Description "git annotated tag" -Command { git tag -a $tag -m "Release $tag" }
    Invoke-CheckedCommand -Description "git push main and tag" -Command { git push origin main --follow-tags }

    $remoteUrl = ((Invoke-CheckedCommand -Description "git remote lookup" -Command { git remote get-url origin }) | Select-Object -First 1).Trim()
    $repoWebUrl = Convert-RemoteUrlToWebUrl -RemoteUrl $remoteUrl

    Write-Host "`nRelease push complete." -ForegroundColor Green
    Write-Host "Tag: $tag"
    Write-Host "Actions: $repoWebUrl/actions/workflows/release.yml"
    Write-Host "Release: $repoWebUrl/releases/tag/$tag"
} catch {
    if ($versionUpdated -and -not $commitCreated -and $null -ne $originalVersionText) {
        Write-Utf8NoBomFile -Path $versionPath -Value $originalVersionText
        Write-Host "Restored VERSION to $originalVersion" -ForegroundColor Yellow
    }

    Write-Error $_
    exit 1
} finally {
    if ($locationPushed) {
        Pop-Location
    }
}
