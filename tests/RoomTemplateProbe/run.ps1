param(
  [Parameter(Mandatory = $true)][string]$SourceDwg,
  [string]$CoreConsole = 'C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe'
)
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$output = Join-Path $repoRoot 'build/room-template-probe'
$bundleRoot = Join-Path $repoRoot 'build/room-template-bundles'
New-Item -ItemType Directory -Path $output -Force | Out-Null
& dotnet build (Join-Path $repoRoot 'AutoCADCommands/GeneralCommands/AutoCADCommands.GeneralCommands.csproj') -c Release --nologo -v quiet "-p:AutoCADBundleRoot=$bundleRoot" -clp:ErrorsOnly
if ($LASTEXITCODE -ne 0) { throw 'GeneralCommands build failed.' }
& dotnet build (Join-Path $PSScriptRoot 'RoomTemplateProbe.csproj') -c Release -f net8.0-windows --nologo -v quiet
if ($LASTEXITCODE -ne 0) { throw 'Probe build failed.' }
$assemblyPath = Join-Path $repoRoot 'AutoCADCommands/GeneralCommands/bin/Release/net8.0-windows/AutoCADCommands.GeneralCommands.dll'
$probePath = Join-Path $PSScriptRoot 'bin/Release/net8.0-windows/Acies.RoomTemplateProbe.dll'
$previousSource = $env:ACIES_RT_SOURCE
$previousOutput = $env:ACIES_RT_OUTPUT
$previousAssembly = $env:ACIES_RT_ASSEMBLY
$previousCircuitTest = $env:ACIES_RT_CIRCUIT_TEST
try {
  $env:ACIES_RT_SOURCE = (Resolve-Path -LiteralPath $SourceDwg).Path
  $env:ACIES_RT_OUTPUT = $output
  $env:ACIES_RT_ASSEMBLY = $assemblyPath
  $inputCopy = Join-Path $output 'command-input.dwg'
  Copy-Item -LiteralPath $env:ACIES_RT_SOURCE -Destination $inputCopy -Force
  $scripts = @{
    verify = @('FILEDIA', '0', 'SECURELOAD', '0', '_NETLOAD', $probePath, 'RTVERIFY', '_QUIT', '_Y')
    command = @('FILEDIA', '0', 'SECURELOAD', '0', '_NETLOAD', $assemblyPath, '_NETLOAD', $probePath,
      '_UCS', '_World', 'RTCOMMANDSETUP', 'RTP', 'RT_PROBE', '0,0', 'Back', '100,100', '200,200',
      'RTCOMMANDASSERT', '_QUIT', '_Y')
    view = @('FILEDIA', '0', 'SECURELOAD', '0', '_NETLOAD', $probePath, 'RTVIEWPROBE', '_QUIT', '_Y')
  }
  foreach ($circuitChoice in @('skip', 'yes', 'default')) {
    $choiceInput = switch ($circuitChoice) { 'skip' { 'No' }; 'yes' { 'Yes' }; 'default' { '' } }
    $scripts["circuit-$circuitChoice"] = @('FILEDIA', '0', 'SECURELOAD', '0', '_NETLOAD', $assemblyPath, '_NETLOAD', $probePath,
      '_UCS', '_World', 'RTCOMMANDSETUP', 'RTP', 'RT_PROBE', '0,0', 'Back', '100,100', '200,200',
      $choiceInput, 'RTCOMMANDASSERT', '_QUIT', '_Y')
  }
  foreach ($scenario in @('verify', 'view', 'command', 'circuit-skip', 'circuit-yes', 'circuit-default')) {
    $env:ACIES_RT_CIRCUIT_TEST = if ($scenario.StartsWith('circuit-')) { $scenario.Substring(8) } else { $null }
    $scriptPath = Join-Path $output "$scenario.scr"
    $scripts[$scenario] | ForEach-Object { $_ -replace '\\', '/' } | Set-Content -LiteralPath $scriptPath -Encoding ascii
    $reportPath = Join-Path $output "$scenario.txt"
    # Avoid accepting an old success report if a new Core Console run fails.
    if (Test-Path -LiteralPath $reportPath) { Remove-Item -LiteralPath $reportPath }
    $arguments = '/s "' + $scriptPath + '" /l en-US'
    if ($scenario -ne 'verify') { $arguments = '/i "' + $inputCopy + '" ' + $arguments }
    $process = Start-Process -FilePath $CoreConsole -ArgumentList $arguments -WorkingDirectory $output -WindowStyle Hidden `
      -RedirectStandardOutput (Join-Path $output "$scenario.log") -RedirectStandardError (Join-Path $output "$scenario.err") -PassThru
    if (-not $process.WaitForExit(45000)) { $process.Kill(); throw "$scenario timed out." }
    if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $reportPath)) { throw "$scenario failed to produce a report." }
    $report = Get-Content -LiteralPath $reportPath -Raw
    Write-Output $report
    if ($report.Contains('FAIL:') -or -not $report.Contains('PASS:')) { throw "$scenario assertions failed." }
    if ($scenario.StartsWith('circuit-')) {
      $log = Get-Content -LiteralPath (Join-Path $output "$scenario.log") -Encoding Unicode -Raw
      if (-not $log.Contains('Name and circuit this room now?')) { throw 'Missing final circuiting choice.' }
      if ($scenario -eq 'circuit-skip') {
        if (-not $log.Contains('Room naming and circuiting skipped.') -or $log.Contains('RTP requires a panel name')) { throw 'Skip did not bypass panel preparation.' }
      } elseif (-not $log.Contains('RTP requires a panel name')) { throw 'Yes/default did not enter circuiting preparation.' }
      if ($log.Contains('Enter room name') -or $log.Contains('Select the panel schedule now.')) { throw 'Unexpected naming or schedule prompt in the missing-panel test.' }
      Write-Output "PASS: $scenario preserves placement and follows the expected circuiting path."
    }
  }
} finally {
  $env:ACIES_RT_SOURCE = $previousSource
  $env:ACIES_RT_OUTPUT = $previousOutput
  $env:ACIES_RT_ASSEMBLY = $previousAssembly
  $env:ACIES_RT_CIRCUIT_TEST = $previousCircuitTest
}
