param(
  [string]$AcadCore
)

# --- SCRIPT CONFIGURATION ---
# Logic to find AutoCAD Core Console
if ($AcadCore -and (Test-Path -Path $AcadCore)) {
  $acadCore = $AcadCore
  Write-Host "PROGRESS: Using specified AutoCAD Core Console: $AcadCore"
}
else {
  $acadCore = $null
  $years = 2025, 2024, 2023, 2022, 2021, 2020

  foreach ($year in $years) {
    $possiblePath = "C:\Program Files\Autodesk\AutoCAD $year\accoreconsole.exe"
    if (Test-Path -Path $possiblePath) {
      $acadCore = $possiblePath
      Write-Host "PROGRESS: Found AutoCAD $year Core Console."
      break # Stop searching once the latest version is found
    }
  }
}

# Name for the combined PDF file
$combinedPdfName = "combined.pdf"
# Name of the Python executable (can be python, python3, or a full path)
$pythonExecutable = "python"

# --- DEFINE AVAILABLE PAPER SIZES ---
$paperSizes = @(
  "ARCH full bleed E1 (30.00 x 42.00 Inches)",
  "ARCH full bleed D (24.00 x 36.00 Inches)",
  "ANSI full bleed D (22.00 x 34.00 Inches)"
)

# --- SCRIPT AND PYTHON VALIDATION ---
# Get the directory where this script is located
$scriptRoot = $PSScriptRoot
$pythonScriptPath = Join-Path $scriptRoot "merge_pdfs.py"
# Check if the Python script exists
if (-not (Test-Path $pythonScriptPath)) {
  Write-Host "PROGRESS: ERROR: 'merge_pdfs.py' not found."
  exit 1
}
# Check if Python is available in the system's PATH
$pythonCheck = Get-Command $pythonExecutable -ErrorAction SilentlyContinue
if (-not $pythonCheck) {
  Write-Host "PROGRESS: ERROR: Python executable ('$pythonExecutable') not found in PATH."
  exit 1
}
# Relaunch in STA mode for the file picker dialog to work correctly
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $ps = (Get-Process -Id $PID).Path
  Start-Process -FilePath $ps -ArgumentList @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") -Wait
  exit
}
# Validate that accoreconsole.exe exists
if ([string]::IsNullOrEmpty($acadCore) -or -not (Test-Path $acadCore)) {
  Write-Host "PROGRESS: ERROR: AutoCAD Core Console not found for versions 2020-2025."
  Write-Host "Please ensure AutoCAD is installed in the default 'C:\Program Files\Autodesk' directory."
  exit 1
}

# --- Let the user select a paper size ---
Write-Host "PROGRESS: Waiting for user input..."
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Paper Size"
$form.Size = New-Object System.Drawing.Size(400, 150)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(280, 20)
$label.Text = "Please select a paper size for plotting:"
$form.Controls.Add($label)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 40)
$comboBox.Size = New-Object System.Drawing.Size(360, 20)
$comboBox.DropDownStyle = "DropDownList"
$paperSizes | ForEach-Object { [void] $comboBox.Items.Add($_) }
$comboBox.SelectedIndex = 0
$form.Controls.Add($comboBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(150, 70)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
  Write-Host "PROGRESS: ERROR: Operation cancelled by user."; exit
}
$selectedPaperSize = $comboBox.SelectedItem

# --- Let the user select DWGs via a file explorer dialog ---
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "Select DWG file(s) to plot"
$dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*"
$dlg.Multiselect = $true
$dlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dlg.FileNames) {
  Write-Host "PROGRESS: ERROR: No files selected."; exit
}
$files = $dlg.FileNames

# --- BATCH PROCESSING SETUP ---
Write-Host "PROGRESS: Preparing to plot $($files.Count) file(s)..."
$basePlotDir = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "AutoCAD Plots"
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$firstFileParentDirName = (Get-Item -Path $files[0]).Directory.Parent.Name
$batchOutputDir = Join-Path -Path (Join-Path -Path $basePlotDir -ChildPath $firstFileParentDirName) -ChildPath $timestamp
New-Item -ItemType Directory -Force -Path $batchOutputDir | Out-Null

# --- SINGLE LOG FILE SETUP ---
$logFile = Join-Path $batchOutputDir "_BatchPlotLog.txt"
"===== Batch Plot Started: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') =====" | Out-File $logFile
"Selected Paper Size: $selectedPaperSize" | Out-File $logFile -Append
"AutoCAD Core Used: $acadCore" | Out-File $logFile -Append
"Output Folder: $batchOutputDir" | Out-File $logFile -Append
"Processing $($files.Count) files..." | Out-File $logFile -Append

# Initialize lists to track progress and generated files
$allGeneratedPdfs = [System.Collections.ArrayList]::new()
$failed = @()
$i = 0

# --- Main Processing Loop (Plotting ONLY) ---
foreach ($dwgPath in $files) {
  $i++
  $dwgItem = Get-Item $dwgPath
  $dwgNameWithoutExt = $dwgItem.BaseName
  Write-Host "PROGRESS: Plotting $i of $($files.Count): $($dwgItem.Name)"
    
  "===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Start Plotting: $($dwgItem.Name) =====" | Out-File $logFile -Append
    
  # Create AutoLISP file for plotting
  $lispFile = Join-Path $env:TEMP "plot_layouts.lsp"
  $lispOutputDir = $batchOutputDir -replace '\\', '\\'
  $lispContent = @"
(defun UpdateBlockOLELinks (blk / res)
  (vlax-for obj blk
    (if (equal (vla-get-ObjectName obj) "AcDbOle2Frame")
      (progn
        (setq res (vl-catch-all-apply 'vla-Update (list obj)))
        (if (vl-catch-all-error-p res)
          (princ (strcat "\nOLE update failed: " (vl-catch-all-error-message res)))
          (setq *ole-updated-count* (1+ *ole-updated-count*))
        )
      )
    )
  )
)

(defun RefreshLinkedOLEs (/ acad doc)
  (vl-load-com)
  (setq acad (vlax-get-acad-object))
  (setq doc (vla-get-ActiveDocument acad))
  (setq *ole-updated-count* 0)
  (UpdateBlockOLELinks (vla-get-ModelSpace doc))
  (vlax-for lay (vla-get-Layouts doc)
    (if (/= (strcase (vla-get-Name lay)) "MODEL")
      (UpdateBlockOLELinks (vla-get-Block lay))
    )
  )
  (princ (strcat "\nOLE links refreshed: " (itoa *ole-updated-count*)))
)

(defun SafeRefreshLinkedOLEs (/ res)
  (setq res (vl-catch-all-apply 'RefreshLinkedOLEs '()))
  (if (vl-catch-all-error-p res)
    (princ (strcat "\nOLE refresh skipped: " (vl-catch-all-error-message res)))
  )
)

(defun c:PlotAllLayouts ()
  (setvar "BACKGROUNDPLOT" 0)
  (setvar "FILEDIA" 0)
  (SafeRefreshLinkedOLEs)
  (setq main-dict (namedobjdict))
  (setq layout-dict (dictsearch main-dict "ACAD_LAYOUT"))
  (foreach item layout-dict
    (if (= (car item) 3)
      (progn
        (setq layout-name (cdr item))
        (if (/= (strcase layout-name) "MODEL")
          (progn
            (setvar 'CTAB layout-name)
            (setq pdfName (strcat "$lispOutputDir\\" "$dwgNameWithoutExt" "-" layout-name ".pdf"))
            (command "-PLOT" "Y" "" "DWG to PDF.pc3" "$selectedPaperSize" "I" "L" "N" "L" "1:1" "0.00,0.00" "Y" "510-monochrome.ctb" "Y" "N" "N" "N" pdfName "N" "Y")
          )
        )
      )
    )
  )
  (command "QUIT" "Y")
  (princ)
)
"@
  Set-Content -Encoding ASCII -Path $lispFile -Value $lispContent
    
  $scriptFile = Join-Path $env:TEMP "run_plot.scr"
  $lispPathForScript = $lispFile -replace '\\', '/'
  $scriptContent = @"
(load "$lispPathForScript")
PlotAllLayouts
"@
  Set-Content -Encoding ASCII -Path $scriptFile -Value $scriptContent
    
  & $acadCore /i "$dwgPath" /s "$scriptFile" 2>&1 | Tee-Object -FilePath $logFile -Append
  $code = $LASTEXITCODE
  "ExitCode: $code" | Out-File $logFile -Append
    
  if ($code -ne 0) {
    $failed += $dwgPath
  }
  else {
    $individualPdfs = @(Get-ChildItem -Path $batchOutputDir -Filter "$($dwgNameWithoutExt)-*.pdf")
    if ($individualPdfs.Count -gt 0) {
      $pdfPaths = @($individualPdfs | ForEach-Object { $_.FullName } | Sort-Object)
      $allGeneratedPdfs.AddRange($pdfPaths)
    }
  }
  "===== $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') Done: $($dwgItem.Name) =====" | Out-File $logFile -Append
}

# --- FINAL MERGE (After the loop) ---
if ($allGeneratedPdfs.Count -gt 0) {
  Write-Host "PROGRESS: Combining $($allGeneratedPdfs.Count) generated PDFs..."
  $combinedPdfPath = Join-Path $batchOutputDir $combinedPdfName
   
  $pythonArgs = @(
    "`"$combinedPdfPath`""
  ) + $allGeneratedPdfs.ForEach({ "`"$_`"" })
    
  $mergeOutput = & $pythonExecutable $pythonScriptPath $pythonArgs 2>&1
   
  if ($LASTEXITCODE -eq 0) {
    "Python script executed successfully." | Out-File $logFile -Append
    "$mergeOutput" | Out-File $logFile -Append
  }
  else {
    Write-Host "PROGRESS: ERROR: PDF merging failed."
    "PDF Merging Failed. Output from Python script:" | Out-File $logFile -Append
    "$mergeOutput" | Out-File $logFile -Append
  }
}
else {
  Write-Host "PROGRESS: No PDFs were generated to combine."
  "No PDFs were generated to merge." | Out-File $logFile -Append
}

# --- CLEANUP & NOTIFICATION ---
"===== Batch Plot Finished: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') =====" | Out-File $logFile -Append
if ($failed.Count) {
  $failMsg = "One or more files failed to plot: $($failed -join ', ')"
  Write-Host "PROGRESS: ERROR: $failMsg"
  $failMsg | Out-File $logFile -Append
}

# Open the final output folder as the very last step
Invoke-Item $batchOutputDir