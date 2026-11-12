param(
  [string]$DotnetPath = 'dotnet',
  [string]$CoreConsole = 'C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe'
)
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$output = Join-Path $repoRoot 'build/txtl-probe'
$bundleRoot = Join-Path $repoRoot 'build/txtl-bundles'
New-Item -ItemType Directory -Path $output -Force | Out-Null
& $DotnetPath build (Join-Path $repoRoot 'AutoCADCommands/GeneralCommands/AutoCADCommands.GeneralCommands.csproj') -c Release --nologo -v quiet "-p:AutoCADBundleRoot=$bundleRoot" -clp:ErrorsOnly
if ($LASTEXITCODE -ne 0) { throw 'GeneralCommands build failed.' }
& $DotnetPath build (Join-Path $PSScriptRoot 'TxtlProbe.csproj') -c Release '-p:TargetFrameworks=net8.0-windows' --nologo -v quiet -clp:ErrorsOnly
if ($LASTEXITCODE -ne 0) { throw 'Probe build failed.' }
$assembly = (Join-Path $bundleRoot 'ElectricalCommands.GeneralCommands.bundle/Contents/net8.0-windows/AutoCADCommands.GeneralCommands.dll').Replace('\', '/')
$probe = (Join-Path $PSScriptRoot 'bin/Release/net8.0-windows/Acies.TxtlProbe.dll').Replace('\', '/')
$savedAssembly = $env:ACIES_TXTL_ASSEMBLY
$savedReport = $env:ACIES_TXTL_REPORT
try {
  $env:ACIES_TXTL_ASSEMBLY = $assembly
  $scenarios = @('East', 'West', 'North', 'South', 'Circle', 'Ucs', 'EighthScale', 'Level', 'Above', 'OverlapRetry', 'MissingScale', 'Blank', 'Cancel', 'Undo')
  $reportPath = Join-Path $output ('report-' + [guid]::NewGuid().ToString('N') + '.txt')
  $env:ACIES_TXTL_REPORT = $reportPath
  $commands = @('FILEDIA', '0', 'SECURELOAD', '0', '_NETLOAD', ('"' + $assembly + '"'), '_NETLOAD', ('"' + $probe + '"'))
  foreach ($scenario in $scenarios) {
    $commands += @('TXTLFIXTURE', $scenario, 'TXTL')
    if ($scenario -ne 'MissingScale') {
      if ($scenario -eq 'Blank') { $commands += '' }
      else {
        $commands += @('TEST {A} \P', 'Arrow', '0,0')
        if ($scenario -eq 'Cancel') { $commands += '' }
        else {
          $direction = if ($scenario -in @('West', 'North', 'South')) { $scenario } else { 'East' }
          $commands += $direction
          if ($scenario -eq 'Circle') { $commands += 'Circle' }
          if ($scenario -eq 'OverlapRetry') { $commands += '0,0' }
          $commands += switch ($scenario) { 'Level' { '100,0' }; 'Above' { '0,100' }; default { '100,100' } }
          if ($scenario -eq 'Undo') { $commands += '_U' }
        }
      }
    }
    $commands += 'TXTLASSERT'
  }
  $commands += @('_QUIT', '_Y')
  $scriptPath = Join-Path $output 'probe.scr'
  $commands | Set-Content -LiteralPath $scriptPath -Encoding ascii
  $profile = 'TxtlProbe' + [guid]::NewGuid().ToString('N')
  $arguments = '/s "' + $scriptPath + '" /l en-US /isolate ' + $profile + ' "' + (Join-Path $output $profile) + '"'
  $process = Start-Process -FilePath $CoreConsole -ArgumentList $arguments -WorkingDirectory $output -WindowStyle Hidden `
    -RedirectStandardOutput (Join-Path $output 'probe.log') -RedirectStandardError (Join-Path $output 'probe.err') -PassThru
  if (-not $process.WaitForExit(45000)) { $process.Kill(); throw "Probe timed out. See $output" }
  if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $reportPath)) { throw "Probe failed to produce a report. See $output" }
  $report = Get-Content -LiteralPath $reportPath
  Write-Output $report
  if ($report.Count -ne $scenarios.Count -or ($report | Where-Object { -not $_.StartsWith('PASS:') })) { throw 'TXTL assertions failed.' }
} finally {
  $env:ACIES_TXTL_ASSEMBLY = $savedAssembly
  $env:ACIES_TXTL_REPORT = $savedReport
}
