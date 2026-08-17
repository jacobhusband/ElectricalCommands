<#
.SYNOPSIS
Captures diagnostics for mapped or UNC project-folder hangs without changing app or server data.

.DESCRIPTION
By default this script only performs read-only path probes with timeouts, captures SMB mappings,
records relevant Explorer/endpoint event-log entries, and writes JSON/Markdown reports under
Documents\ProjectManagementApp\diagnostics. It only launches Explorer when -OpenExplorer is passed.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DiagnoseProjectPath.ps1 "M:\Client\Project"

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DiagnoseProjectPath.ps1 "M:\Client\Project" -OpenExplorer -ExplorerTarget Both
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [ValidateSet("Input", "Equivalent", "Both")]
    [string]$ExplorerTarget = "Input",

    [switch]$OpenExplorer,

    [int]$SinceMinutes = 120,

    [int]$AfterOpenWaitSeconds = 20,

    [int]$PathProbeTimeoutSeconds = 10,

    [string]$OutputDir = (Join-Path ([Environment]::GetFolderPath("MyDocuments")) "ProjectManagementApp\diagnostics")
)

$ErrorActionPreference = "Stop"

function ConvertTo-NormalizedProjectPath {
    param([string]$RawPath)

    if ([string]::IsNullOrWhiteSpace($RawPath)) {
        $value = ""
    } else {
        $value = $RawPath.Trim().Trim('"').Trim("'")
    }
    if (-not $value) {
        return ""
    }

    $value = $value -replace '/', '\'
    if ($value -match '^[A-Za-z]:$') {
        return "$value\"
    }
    return $value.TrimEnd('\')
}

function Resolve-EquivalentProjectPath {
    param([string]$NormalizedPath)

    if (-not $NormalizedPath) {
        return @{
            Path = ""
            Kind = "unknown"
            Reason = "empty"
        }
    }

    if ($NormalizedPath -match '(?i)^P:\\(.+)$') {
        return @{
            Path = "\\acies.lan\cachedrive\Projects\$($Matches[1])"
            Kind = "unc"
            Reason = "mapped-p-to-unc"
        }
    }

    if ($NormalizedPath -match '(?i)^M:\\(.+)$') {
        return @{
            Path = "\\acies.lan\cachedrive\projects2\$($Matches[1])"
            Kind = "unc"
            Reason = "mapped-m-to-unc"
        }
    }

    if ($NormalizedPath -match '(?i)^\\\\acies\.lan\\cachedrive\\Projects\\(.+)$') {
        return @{
            Path = "P:\$($Matches[1])"
            Kind = "mapped"
            Reason = "unc-projects-to-p"
        }
    }

    if ($NormalizedPath -match '(?i)^\\\\acies\.lan\\cachedrive\\projects2\\(.+)$') {
        return @{
            Path = "M:\$($Matches[1])"
            Kind = "mapped"
            Reason = "unc-projects2-to-m"
        }
    }

    return @{
        Path = ""
        Kind = "unknown"
        Reason = "no-known-equivalent"
    }
}

function Invoke-PathProbeWithTimeout {
    param(
        [string]$Label,
        [string]$CandidatePath,
        [int]$TimeoutSeconds
    )

    $normalized = ConvertTo-NormalizedProjectPath $CandidatePath
    if (-not $normalized) {
        return [ordered]@{
            label = $Label
            path = ""
            exists = $false
            isDirectory = $false
            parentPath = ""
            parentExists = $false
            timedOut = $false
            error = "Path is empty."
        }
    }

    $fallbackParent = ""
    try {
        $fallbackParent = [System.IO.Path]::GetDirectoryName($normalized)
    } catch {
        $fallbackParent = ""
    }

    $probeCommand = @'
$ErrorActionPreference = "Stop"
$candidate = $env:DIAG_PROJECT_PATH
$label = $env:DIAG_PROJECT_LABEL
$parent = ""
try {
    $parent = [System.IO.Path]::GetDirectoryName($candidate)
} catch {
    $parent = ""
}

$exists = $false
$isDirectory = $false
$parentExists = $false
$errorMessage = ""

try {
    $exists = Test-Path -LiteralPath $candidate
    if ($exists) {
        $isDirectory = Test-Path -LiteralPath $candidate -PathType Container
    }
    if ($parent) {
        $parentExists = Test-Path -LiteralPath $parent -PathType Container
    }
} catch {
    $errorMessage = $_.Exception.Message
}

[ordered]@{
    label = $label
    path = $candidate
    exists = [bool]$exists
    isDirectory = [bool]$isDirectory
    parentPath = $parent
    parentExists = [bool]$parentExists
    timedOut = $false
    error = $errorMessage
} | ConvertTo-Json -Compress
'@

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $encodedProbeCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probeCommand))
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encodedProbeCommand"
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.EnvironmentVariables["DIAG_PROJECT_PATH"] = $normalized
    $psi.EnvironmentVariables["DIAG_PROJECT_LABEL"] = $Label

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    try {
        [void]$process.Start()

        $timeoutMs = [Math]::Max(1, $TimeoutSeconds) * 1000
        if (-not $process.WaitForExit($timeoutMs)) {
            try {
                $process.Kill()
            } catch {
            }
            return [ordered]@{
                label = $Label
                path = $normalized
                exists = $false
                isDirectory = $false
                parentPath = $fallbackParent
                parentExists = $false
                timedOut = $true
                error = "Path probe timed out after $TimeoutSeconds seconds."
            }
        }

        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        if ($process.ExitCode -ne 0 -or [string]::IsNullOrWhiteSpace($stdout)) {
            $message = $stderr.Trim()
            if (-not $message) {
                $message = "Path probe exited with code $($process.ExitCode)."
            }
            return [ordered]@{
                label = $Label
                path = $normalized
                exists = $false
                isDirectory = $false
                parentPath = $fallbackParent
                parentExists = $false
                timedOut = $false
                error = $message
            }
        }

        $payload = $stdout | ConvertFrom-Json
        return [ordered]@{
            label = [string]$payload.label
            path = [string]$payload.path
            exists = [bool]$payload.exists
            isDirectory = [bool]$payload.isDirectory
            parentPath = [string]$payload.parentPath
            parentExists = [bool]$payload.parentExists
            timedOut = [bool]$payload.timedOut
            error = [string]$payload.error
        }
    } catch {
        return [ordered]@{
            label = $Label
            path = $normalized
            exists = $false
            isDirectory = $false
            parentPath = $fallbackParent
            parentExists = $false
            timedOut = $false
            error = $_.Exception.Message
        }
    } finally {
        if ($process -and -not $process.HasExited) {
            try {
                $process.Kill()
            } catch {
            }
        }
        if ($process) {
            $process.Dispose()
        }
    }
}

function Test-ProjectPathState {
    param(
        [string]$Label,
        [string]$CandidatePath
    )

    return Invoke-PathProbeWithTimeout `
        -Label $Label `
        -CandidatePath $CandidatePath `
        -TimeoutSeconds $PathProbeTimeoutSeconds
}

function Get-RelevantApplicationEvents {
    param(
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    try {
        Get-WinEvent -FilterHashtable @{
            LogName = "Application"
            StartTime = $StartTime
            EndTime = $EndTime
        } -ErrorAction Stop |
            Where-Object {
                $_.ProviderName -match 'Application Error|Application Hang|Windows Error Reporting' -and
                $_.Message -match 'explorer\.exe|AEMAgent\.exe|svchost\.exe_DsmSvc|svchost\.exe|ACIES Scheduler|msedgewebview2|pywebview'
            } |
            Select-Object TimeCreated, ProviderName, Id, LevelDisplayName,
                @{Name = "Message"; Expression = { ($_.Message -replace "`r?`n", " ") }}
    } catch {
        @([pscustomobject]@{
            TimeCreated = $null
            ProviderName = "diagnostic"
            Id = 0
            LevelDisplayName = "Error"
            Message = "Could not read Application event log: $($_.Exception.Message)"
        })
    }
}

function Get-RelevantSystemEvents {
    param(
        [datetime]$StartTime,
        [datetime]$EndTime
    )

    try {
        Get-WinEvent -FilterHashtable @{
            LogName = "System"
            StartTime = $StartTime
            EndTime = $EndTime
        } -ErrorAction Stop |
            Where-Object {
                $_.ProviderName -match 'LanmanWorkstation|MRxSmb|SMBClient|Tcpip|NETLOGON|GroupPolicy|RasMan|DistributedCOM' -or
                $_.Message -match 'SMB|network path|server|redirector|connection|mapped drive|VPN|Deployed Printer Connections'
            } |
            Select-Object TimeCreated, ProviderName, Id, LevelDisplayName,
                @{Name = "Message"; Expression = { ($_.Message -replace "`r?`n", " ") }}
    } catch {
        @([pscustomobject]@{
            TimeCreated = $null
            ProviderName = "diagnostic"
            Id = 0
            LevelDisplayName = "Error"
            Message = "Could not read System event log: $($_.Exception.Message)"
        })
    }
}

function Get-SmbMappingSnapshot {
    try {
        Get-SmbMapping -ErrorAction Stop |
            Select-Object LocalPath, RemotePath, Status, UserName
    } catch {
        @([pscustomobject]@{
            LocalPath = ""
            RemotePath = ""
            Status = "unavailable"
            UserName = ""
        })
    }
}

function Get-EndpointSnapshot {
    $services = Get-Service -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match 'DsmSvc|CagService|AEM|Centra|Datto|HUNTAgent|dattorollbackservice|LanmanWorkstation|WebClient|WinDefend|Sense' -or
            $_.DisplayName -match 'Device Setup|Centra|Datto|Managed|Endpoint|Defender|Workstation|WebClient'
        } |
        Select-Object Name, DisplayName, Status, StartType

    $processes = Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ProcessName -match 'ACIES|Scheduler|python|pywebview|msedgewebview2|explorer|AEMAgent|CagService|Centra|Datto|HUNT|DsmSvc'
        } |
        Select-Object ProcessName, Id, CPU, StartTime, Path

    return [ordered]@{
        services = @($services)
        processes = @($processes)
    }
}

function Start-ExplorerProbe {
    param([string]$TargetPath)

    $normalized = ConvertTo-NormalizedProjectPath $TargetPath
    if (-not $normalized) {
        return [ordered]@{
            path = ""
            status = "skipped"
            message = "No path provided."
        }
    }

    try {
        $startedAt = Get-Date
        $process = Start-Process -FilePath "explorer.exe" -ArgumentList @($normalized) -PassThru
        return [ordered]@{
            path = $normalized
            status = "launched"
            processId = $process.Id
            startedAt = $startedAt.ToString("o")
            message = "Explorer launch requested. Check post-launch events for hangs."
        }
    } catch {
        return [ordered]@{
            path = $normalized
            status = "error"
            message = $_.Exception.Message
        }
    }
}

function New-DiagnosticAssessment {
    param(
        [object]$InputState,
        [object]$EquivalentState,
        [object[]]$PostOpenEvents,
        [object[]]$SmbMappings,
        [object]$EndpointSnapshot
    )

    $findings = New-Object System.Collections.Generic.List[string]

    if ($InputState.timedOut -and $EquivalentState.path -and $EquivalentState.timedOut) {
        $findings.Add("mapped-and-unc-path-probes-timed-out")
    } elseif ($InputState.timedOut) {
        $findings.Add("input-path-probe-timed-out")
    } elseif ($EquivalentState.path -and $EquivalentState.timedOut) {
        $findings.Add("equivalent-path-probe-timed-out")
    }

    if (-not $InputState.exists -and $InputState.parentExists) {
        $findings.Add("input-path-missing-parent-exists")
    }

    if ($EquivalentState.path -and -not $EquivalentState.exists -and $EquivalentState.parentExists) {
        $findings.Add("equivalent-path-missing-parent-exists")
    }

    if ($InputState.path -and $EquivalentState.path) {
        if ($InputState.exists -and -not $EquivalentState.exists) {
            $findings.Add("input-exists-equivalent-missing")
        } elseif (-not $InputState.exists -and $EquivalentState.exists) {
            $findings.Add("input-missing-equivalent-exists")
        }
    }

    $explorerHang = @($PostOpenEvents | Where-Object { $_.Message -match 'AppHangB1|explorer\.exe.*stopped interacting|P1:\s*explorer\.exe' })
    if ($explorerHang.Count -gt 0) {
        $findings.Add("explorer-hang-events-present")
    }

    $endpointEvents = @($PostOpenEvents | Where-Object { $_.Message -match 'AEMAgent\.exe|svchost\.exe_DsmSvc|Datto|CentraStage' })
    if ($endpointEvents.Count -gt 0) {
        $findings.Add("endpoint-or-dsmsvc-events-present")
    }

    $notOkMappings = @($SmbMappings | Where-Object { $_.Status -and $_.Status -ne "OK" })
    if ($notOkMappings.Count -gt 0) {
        $findings.Add("smb-mapping-not-ok")
    }

    $endpointServices = @($EndpointSnapshot.services | Where-Object { $_.Name -match 'CagService|HUNTAgent|dattorollbackservice' -or $_.DisplayName -match 'Datto' })
    if ($endpointServices.Count -gt 0) {
        $findings.Add("datto-services-present")
    }

    if ($findings.Count -eq 0) {
        $findings.Add("no-obvious-local-signal")
    }

    return @($findings)
}

function Write-MarkdownReport {
    param(
        [string]$ReportPath,
        [object]$Report
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# Project Path Explorer Diagnostic")
    $lines.Add("")
    $lines.Add("- Created: $($Report.createdAt)")
    $lines.Add("- Input path: ``$($Report.input.path)``")
    $lines.Add("- Equivalent path: ``$($Report.equivalent.path)``")
    $lines.Add("- Explorer probe: $($Report.explorerProbe.enabled)")
    $lines.Add("")
    $lines.Add("## Path Checks")
    foreach ($state in @($Report.input, $Report.equivalent)) {
        if (-not $state.path) {
            continue
        }
        $lines.Add("- $($state.label): exists=$($state.exists), isDirectory=$($state.isDirectory), parentExists=$($state.parentExists), timedOut=$($state.timedOut)")
        $lines.Add("  - Path: ``$($state.path)``")
        $lines.Add("  - Parent: ``$($state.parentPath)``")
        if ($state.error) {
            $lines.Add("  - Error: $($state.error)")
        }
    }
    $lines.Add("")
    $lines.Add("## Assessment Signals")
    foreach ($finding in @($Report.assessmentFindings)) {
        $lines.Add("- $finding")
    }
    $lines.Add("")
    $lines.Add("## SMB Mappings")
    foreach ($mapping in @($Report.smbMappings)) {
        $lines.Add("- $($mapping.LocalPath) -> $($mapping.RemotePath) [$($mapping.Status)]")
    }
    $lines.Add("")
    $lines.Add("## Explorer Probe Results")
    foreach ($probe in @($Report.explorerProbe.results)) {
        $lines.Add("- $($probe.status): ``$($probe.path)`` $($probe.message)")
    }
    $lines.Add("")
    $lines.Add("## Relevant Application Events")
    foreach ($event in @($Report.events.application | Select-Object -First 25)) {
        $lines.Add("- $($event.TimeCreated) [$($event.ProviderName) $($event.Id)] $($event.Message)")
    }
    $lines.Add("")
    $lines.Add("## Relevant System Events")
    foreach ($event in @($Report.events.system | Select-Object -First 25)) {
        $lines.Add("- $($event.TimeCreated) [$($event.ProviderName) $($event.Id)] $($event.Message)")
    }

    Set-Content -LiteralPath $ReportPath -Value $lines -Encoding UTF8
}

$createdAt = Get-Date
$eventStart = $createdAt.AddMinutes(-1 * [Math]::Max(1, $SinceMinutes))
$normalizedInput = ConvertTo-NormalizedProjectPath $Path
$equivalent = Resolve-EquivalentProjectPath $normalizedInput

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$inputState = Test-ProjectPathState -Label "input" -CandidatePath $normalizedInput
$equivalentState = Test-ProjectPathState -Label "equivalent" -CandidatePath $equivalent.Path
$smbMappings = @(Get-SmbMappingSnapshot)
$endpointSnapshot = Get-EndpointSnapshot
$applicationEventsBefore = @(Get-RelevantApplicationEvents -StartTime $eventStart -EndTime $createdAt)
$systemEventsBefore = @(Get-RelevantSystemEvents -StartTime $eventStart -EndTime $createdAt)

$explorerResults = @()
$openStartedAt = $null
if ($OpenExplorer) {
    $openStartedAt = Get-Date
    if ($ExplorerTarget -eq "Input" -or $ExplorerTarget -eq "Both") {
        $explorerResults += Start-ExplorerProbe -TargetPath $normalizedInput
    }
    if (($ExplorerTarget -eq "Equivalent" -or $ExplorerTarget -eq "Both") -and $equivalent.Path) {
        $explorerResults += Start-ExplorerProbe -TargetPath $equivalent.Path
    }
    Start-Sleep -Seconds ([Math]::Max(0, $AfterOpenWaitSeconds))
}

$capturedAt = Get-Date
$applicationEventsAfter = @(Get-RelevantApplicationEvents -StartTime $eventStart -EndTime $capturedAt)
$systemEventsAfter = @(Get-RelevantSystemEvents -StartTime $eventStart -EndTime $capturedAt)
$postOpenEvents = if ($openStartedAt) {
    @($applicationEventsAfter | Where-Object { $_.TimeCreated -and $_.TimeCreated -ge $openStartedAt })
} else {
    @()
}

$report = [ordered]@{
    createdAt = $createdAt.ToString("o")
    capturedAt = $capturedAt.ToString("o")
    input = $inputState
    equivalent = $equivalentState
    equivalentResolution = $equivalent
    explorerProbe = [ordered]@{
        enabled = [bool]$OpenExplorer
        target = $ExplorerTarget
        startedAt = if ($openStartedAt) { $openStartedAt.ToString("o") } else { "" }
        waitSeconds = $AfterOpenWaitSeconds
        results = @($explorerResults)
    }
    pathProbeTimeoutSeconds = $PathProbeTimeoutSeconds
    smbMappings = @($smbMappings)
    endpoint = $endpointSnapshot
    events = [ordered]@{
        application = @($applicationEventsAfter)
        system = @($systemEventsAfter)
        applicationBeforeExplorerProbe = @($applicationEventsBefore)
        systemBeforeExplorerProbe = @($systemEventsBefore)
        applicationAfterExplorerProbe = @($postOpenEvents)
    }
}

$report.assessmentFindings = New-DiagnosticAssessment `
    -InputState $inputState `
    -EquivalentState $equivalentState `
    -PostOpenEvents $postOpenEvents `
    -SmbMappings $smbMappings `
    -EndpointSnapshot $endpointSnapshot

$stamp = $createdAt.ToString("yyyyMMdd-HHmmss")
$safeName = ($normalizedInput -replace '^[A-Za-z]:\\', '' -replace '^\\\\', '' -replace '[^A-Za-z0-9._-]+', '_').Trim('_')
if (-not $safeName) {
    $safeName = "project-path"
}
if ($safeName.Length -gt 80) {
    $safeName = $safeName.Substring(0, 80)
}

$jsonPath = Join-Path $OutputDir "project-path-diagnostic-$stamp-$safeName.json"
$mdPath = Join-Path $OutputDir "project-path-diagnostic-$stamp-$safeName.md"

$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
Write-MarkdownReport -ReportPath $mdPath -Report $report

[pscustomobject]@{
    Status = "success"
    JsonReport = $jsonPath
    MarkdownReport = $mdPath
    AssessmentFindings = $report.assessmentFindings -join ", "
    InputExists = $inputState.exists
    EquivalentExists = $equivalentState.exists
    ExplorerProbeRan = [bool]$OpenExplorer
}
