# removeXrefPaths.ps1
# Select DWGs -> accoreconsole loads your .NET DLL via NETLOAD -> runs STRIPREFPATHS
# Logs to Documents\StripRefPathsLogs\strip_xref_paths_YYYYMMDD_HHMMSS.log

$acadCore = "C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe"
$dll      = "C:\Users\JacobH\Dotnet\Finalize\StripRefPaths\bin\Release\net8.0-windows\StripRefPaths.dll"

# Relaunch in STA for file picker
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $ps = (Get-Process -Id $PID).Path
  Start-Process -FilePath $ps -ArgumentList @("-NoProfile","-STA","-ExecutionPolicy","Bypass","-File","`"$PSCommandPath`"") -Wait
  exit
}

if (-not (Test-Path $acadCore)) { Write-Host "accoreconsole not found: $acadCore"; Read-Host "Press Enter"; exit 1 }
if (-not (Test-Path $dll))      { Write-Host "Plugin DLL not found: $dll"; Read-Host "Press Enter"; exit 1 }

# Stage AutoLISP CLEANUP command so each drawing is tidied after ref paths are stripped
$lisp = Join-Path $env:TEMP "cleanup_tmp.lsp"
$lispContent = @'
(defun c:CLEANUP (/ oldCMDECHO oldTab)
  ;; Remember current command echo setting, then turn it off:
  (setq oldCMDECHO (getvar "CMDECHO")
        oldTab     (getvar "CTAB"))
  (setvar "CMDECHO" 0)

  ;; Ensure commands run from Model space
  (if (/= oldTab "Model")
    (command "_.MODEL")
  )

  ;; Use SETBYLAYER on everything in the current space
  ;;   "Y" => Change ByBlock to ByLayer?
  ;;   "Y" => Include blocks?
  (command
    "_.-SETBYLAYER" "All" "" "Y" "Y")

  ;; Purge the entire drawing; "All", "*" (everything), "No" to confirm each
  (command "_.-PURGE" "All" "*" "N")

  ;; Audit the drawing; "Yes" to fix errors
  (command "_.AUDIT" "Y")

  ;; Restore layout/model state if it changed
  (if (/= oldTab (getvar "CTAB"))
    (if (= oldTab "Model")
      (command "_.MODEL")
      (command "_.LAYOUT" "Set" oldTab)
    )
  )

  ;; Restore command echo setting
  (setvar "CMDECHO" oldCMDECHO)
  (princ)
)
'@
Set-Content -Encoding ASCII -Path $lisp -Value $lispContent

# Make .scr that NETLOADs the DLL, strips refs, and then runs CLEANUP
$script = Join-Path $env:TEMP "run_STRIP.scr"
$lispPathForScript = ($lisp -replace '\\', '/')
$scriptContent = @"
CMDECHO 1
FILEDIA 0
SECURELOAD 0
NETLOAD
"$dll"
STRIPREFPATHS
(load "$lispPathForScript")
CLEANUP
QSAVE
QUIT
"@
Set-Content -Encoding ASCII -Path $script -Value $scriptContent

# Prepare log
$logDir = Join-Path $env:USERPROFILE "Documents\StripRefPathsLogs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$log = Join-Path $logDir ("strip_xref_paths_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
"===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Start =====" | Out-File $log
"AutoCAD Core: $acadCore"  | Out-File $log -Append
"Plugin DLL : $dll"        | Out-File $log -Append
"Cleanup LISP: $lisp"      | Out-File $log -Append
"Script     : $script"     | Out-File $log -Append

# Pick DWGs
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select DWG file(s) to strip saved XREF/underlay/image paths"
$dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*"
$dlg.Multiselect = $true
$dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dlg.FileNames) {
  "Cancelled by user." | Out-File $log -Append
  Write-Host "Cancelled."; exit
}
$files = $dlg.FileNames
"Selected $($files.Count) file(s)." | Out-File $log -Append

$failed = @()
$i = 0
foreach ($dwg in $files) {
  $i++
  $name = [IO.Path]::GetFileName($dwg)
  Write-Progress -Activity "Processing DWGs" -Status "$i of $($files.Count): $name" -PercentComplete ([int](($i/$files.Count)*100))
  "----- $(Get-Date -Format 'HH:mm:ss')  $dwg -----" | Out-File $log -Append

  # IMPORTANT: no /ld here; we NETLOAD via the .scr
  & $acadCore /i "$dwg" /s "$script" 2>&1 | Tee-Object -FilePath $log -Append
  $code = $LASTEXITCODE
  "ExitCode: $code" | Out-File $log -Append
  if ($code -ne 0) { $failed += $dwg }
}

Write-Progress -Activity "Processing DWGs" -Completed
"===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Done =====" | Out-File $log -Append
Write-Host "Done. Processed $($files.Count) drawing(s). Log: $log"
Invoke-Item $log

if ($failed.Count) {
  Write-Host "`nOne or more files failed:" -ForegroundColor Yellow
  $failed | ForEach-Object { Write-Host " - $_" }
  Write-Host "`nPress Enter to close..."
  [void](Read-Host)
}



