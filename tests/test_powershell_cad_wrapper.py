import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
MAIN_PY_PATH = REPO_ROOT / "main.py"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
CAD_SCRIPT_PATHS = (
    REPO_ROOT / "scripts" / "PlotDWGs.ps1",
    REPO_ROOT / "scripts" / "ManageLayersDWGs.ps1",
)
CLEAN_XREF_SCRIPT_PATH = REPO_ROOT / "scripts" / "removeXREFPaths.ps1"
DETECT_PDF_SIZE_PATH = REPO_ROOT / "scripts" / "detect_pdf_size.py"


class PowerShellCadWrapperTests(unittest.TestCase):
    def test_publish_supports_arch_full_bleed_e_paper_size(self):
        paper_size = "ARCH full bleed E (36.00 x 48.00 Inches)"
        plot_script = (REPO_ROOT / "scripts" / "PlotDWGs.ps1").read_text(
            encoding="utf-8"
        )
        detector = DETECT_PDF_SIZE_PATH.read_text(encoding="utf-8")

        self.assertIn(paper_size, plot_script)
        self.assertIn(f'"{paper_size}": (36 * 72, 48 * 72)', detector)

    def test_backend_trace_messages_are_logged_but_not_forwarded_to_ui(self):
        text = MAIN_PY_PATH.read_text(encoding="utf-8")
        self.assertIn('if message.startswith("TRACE"):', text)
        self.assertIn("'script_trace'", text)
        self.assertIn("continue", text)
        self.assertIn("def _notify_activity_status(self, payload):", text)
        self.assertIn("window.updateActivityStatus(", text)
        self.assertIn("activity_id=activity_id", text)

    def test_frontend_automation_keeps_cad_trace_helpers_without_workroom_modal(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("async getCadAutoSelectTrace(lineLimit = 200) {", text)
        self.assertIn("async clearCadAutoSelectTrace() {", text)
        self.assertIn("window.pywebview?.api?.get_cad_auto_select_trace", text)
        self.assertIn("window.pywebview?.api?.clear_cad_auto_select_trace", text)
        self.assertNotIn("openWorkroom(projectIndex = 0)", text)
        self.assertNotIn("async runWorkroomTool(toolId) {", text)
        self.assertNotIn("getWorkroomState()", text)
        self.assertNotIn("getToolState(toolId)", text)

    def test_frontend_restores_workroom_launch_context_helpers_for_shared_tools(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function getActiveWorkroomContext() {",
            "function consumePendingCadLaunchContext() {",
            "function buildWorkroomCadLaunchContext() {",
            "function resolveCadLaunchContextForTool() {",
            'source: "workroom",',
            "rootProjectPath,",
            "cadFilePaths: [],",
            "projectName:",
            "deliverableName:",
            "const pendingContext = consumePendingCadLaunchContext();",
            'const checklistModal = document.getElementById("checklistModal");',
            "if (!checklistModal?.open) return null;",
        ):
            self.assertIn(expected, text)

    def test_cad_scripts_use_raw_values_in_sta_relaunch(self):
        for script_path in CAD_SCRIPT_PATHS:
            with self.subTest(script=script_path.name):
                text = script_path.read_text(encoding="utf-8")
                self.assertIn("function Move-FormToPrimaryScreen {", text)
                self.assertIn("[System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea", text)
                self.assertNotIn('`"$PSCommandPath`"', text)
                self.assertNotIn('`"$FilesListPath`"', text)
                self.assertNotIn('`"$AcadCore`"', text)
                self.assertIn('[string]$DefaultDirectory = ""', text)
                self.assertIn(
                    '@("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)',
                    text,
                )
                self.assertIn('@("-FilesListPath", $FilesListPath)', text)
                self.assertIn('@("-DefaultDirectory", $DefaultDirectory)', text)
                self.assertIn("Resolve-DialogInitialDirectory", text)
                self.assertIn(
                    'Write-Host "PROGRESS: Received auto-selected files list: $FilesListPath"',
                    text,
                )
                self.assertIn(
                    'Write-Host "PROGRESS: TRACE files_list_param_bound=$([int]$filesListWasProvided) path=$FilesListPath"',
                    text,
                )
                self.assertIn(
                    'Write-Host "PROGRESS: TRACE branch=auto_selected_files count=$($files.Count)"',
                    text,
                )
                self.assertIn(
                    'Write-Host "PROGRESS: TRACE branch=manual_picker"',
                    text,
                )
                # The picker runs in a hidden-console child process, so an
                # ownerless ShowDialog() lands behind the app window and the
                # tool looks frozen until the app is killed.
                self.assertIn("function Show-DwgFileSelectionPrompt {", text)
                self.assertIn("$promptForm.TopMost = $true", text)
                self.assertIn("$promptForm.ShowInTaskbar = $true", text)
                self.assertIn("$result = $dlg.ShowDialog($promptForm)", text)
                self.assertNotIn("$dlg.ShowDialog()", text)
                self.assertIn('Write-Host "PROGRESS: INPUT_FOLDER: $inputFolder"', text)
                if script_path.name == "PlotDWGs.ps1":
                    self.assertIn('[string]$StripPdfLayers = "true"', text)
                    self.assertIn('$StripPdfLayers = Convert-ToBool $StripPdfLayers $true', text)
                    self.assertIn('$stripPdfLayersScriptPath = Join-Path $scriptRoot "strip_pdf_layers.py"', text)
                    self.assertIn('@("-StripPdfLayers", $StripPdfLayers)', text)
                    self.assertIn('[string]$RefreshExcelOleLinks = "true"', text)
                    self.assertIn(
                        '$RefreshExcelOleLinks = Convert-ToBool $RefreshExcelOleLinks $true',
                        text,
                    )
                    self.assertIn('@("-RefreshExcelOleLinks", $RefreshExcelOleLinks)', text)
                    self.assertIn('Write-Host "PROGRESS: ERROR: \'strip_pdf_layers.py\' not found."', text)
                    self.assertIn('Write-Host "PROGRESS: Removing PDF layers from combined PDF..."', text)
                    self.assertIn('"PDF layer cleanup completed." | Out-File $logFile -Append', text)
                    self.assertIn('Write-Host "PROGRESS: ERROR: PDF layer cleanup failed."', text)
                    self.assertIn('Write-Host "PROGRESS: ERROR: Publish failed because PDF layer cleanup did not complete."', text)
                    self.assertIn('$pdfMergeFailed = $false', text)
                    self.assertIn('Write-Host "PROGRESS: ERROR: PDF merging failed: $mergeErrorSummary"', text)
                    self.assertIn('Write-Host "PROGRESS: ERROR: Publish failed because PDF merge did not complete."', text)
                    self.assertIn('if ($failed.Count -or $pdfLayerCleanupFailed -or $pdfMergeFailed) {', text)
                    self.assertIn('$form.StartPosition = "Manual"', text)
                    self.assertIn('$form.TopMost = $true', text)
                    self.assertIn('$form.ShowInTaskbar = $true', text)
                    self.assertIn(
                        '$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal',
                        text,
                    )
                    self.assertIn("Move-FormToPrimaryScreen $form", text)
                    self.assertIn('$form.BringToFront()', text)
                    self.assertIn(
                        'Write-Host "PROGRESS: Waiting for paper size confirmation..."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: Paper size dialog should be visible on the primary display."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: TRACE branch=paper_size_dialog"',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: COMBINED_PDF: $finalCombinedPdfPath"',
                        text,
                    )
                    self.assertIn('Write-Host "PROGRESS: OUTPUT_FOLDER: $batchOutputDir"', text)
                    self.assertNotIn("Invoke-Item $batchOutputDir", text)
                else:
                    self.assertIn('$form.StartPosition = "Manual"', text)
                    self.assertIn("Move-FormToPrimaryScreen $form", text)
                    self.assertIn(
                        '$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal',
                        text,
                    )
                    self.assertIn('Write-Host "PROGRESS: Reading extracted data..."', text)
                    self.assertIn('$form.TopMost = $true', text)
                    self.assertIn('$form.ShowInTaskbar = $true', text)
                    self.assertIn('$form.BringToFront()', text)
                    self.assertIn(
                        'Write-Host "PROGRESS: Waiting for layer selection..."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: Layer selection dialog should be visible on the primary display."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: TRACE branch=layer_selection_dialog"',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: Processing $fileIndex of $($files.Count): $([IO.Path]::GetFileName($dwgFile)) (attempt $attempt of $MaxUpdateAttemptsPerFile)"',
                        text,
                    )
                    self.assertIn('Write-Host "PROGRESS: OUTPUT_FOLDER: $outputFolder"', text)

    def test_publish_refreshes_linked_excel_ole_through_full_autocad_once(self):
        text = (REPO_ROOT / "scripts" / "PlotDWGs.ps1").read_text(encoding="utf-8")

        for expected in (
            'Join-Path $acadInstallDir "acad.exe"',
            'function Invoke-FullAutoCadOleRefresh {',
            'function Invoke-CorePlotAttempt {',
            '(and typePair (= (cdr typePair) 1))',
            '$oleCheckPresentMarker = "ACIES_OLE_CHECK:PRESENT"',
            '$oleRefreshRequiredMarker = "ACIES_OLE_REFRESH:REQUIRED"',
            '(not *acies-ole-refresh-attempted*)',
            '-OleRefreshAlreadyAttempted $false',
            '-OleRefreshAlreadyAttempted $true',
            '-TimeoutSeconds 180',
            '-WindowStyle Hidden',
            'REGENALL',
            'QSAVE',
            'Publishing cached OLE content.',
            'Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue',
        ):
            self.assertIn(expected, text)

        self.assertIn('Write-Host "PROGRESS: WARNING:', text)
        self.assertNotIn('Write-Host "PROGRESS: ERROR: $($dwgItem.Name) - $($oleRefreshResult.Message)', text)

    def test_publish_excel_ole_refresh_setting_is_bound_in_both_uis(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn("refreshExcelOleLinks: true,", script)
        for control_id in (
            "settings_publish_refreshExcelOleLinks",
            "publish_modal_refreshExcelOleLinks",
        ):
            self.assertIn(control_id, script)
            self.assertIn(f'id="{control_id}"', html)
        self.assertIn("Refresh linked Excel OLE content and save source DWGs", html)

    def test_clean_xrefs_script_reports_output_folder_marker(self):
        text = CLEAN_XREF_SCRIPT_PATH.read_text(encoding="utf-8")
        self.assertIn('Write-Host "PROGRESS: Processing $i of $($sourceItems.Count):', text)
        self.assertIn('Write-Host "PROGRESS: INPUT_FOLDER: $inputFolder"', text)
        self.assertIn('Write-Host "PROGRESS: OUTPUT_FOLDER: $outputFolder"', text)

    def test_publish_combined_pdf_name_uses_first_dwg_folder_context(self):
        text = (REPO_ROOT / "scripts" / "PlotDWGs.ps1").read_text(encoding="utf-8")

        self.assertNotIn('$combinedPdfName = "combined.pdf"', text)
        for expected in (
            "function Resolve-CombinedPdfName {",
            "function Convert-ToSafePdfFileName {",
            "function Get-MaxPdfFileNameLengthForDirectory {",
            "function Get-ToolOutputSummary {",
            "[System.IO.Path]::GetInvalidFileNameChars()",
            '$MaxCombinedPdfFullPathLength = 240',
            '[int]$MaxFileNameLength = 0',
            '[string]$PreserveSuffix = ""',
            '$ellipsis = "..."',
            "$parentDir = $dwgItem.Directory",
            "$projectDir = $parentDir.Parent",
            r"$projectName = ($projectDir.Name -replace '^\s*\d{5,}\s*[-_.]*\s*', '').Trim()",
            '$rawName = "$projectName - $($parentDir.Name)"',
            "$maxFileNameLength = Get-MaxPdfFileNameLengthForDirectory -OutputDirectory $OutputDirectory -MaxFullPathLength $MaxFullPathLength",
            "-PreserveSuffix $parentDir.Name",
            "$combinedPdfName = Resolve-CombinedPdfName -DwgPath ([string]$files[0]) -OutputDirectory $batchOutputDir -MaxFullPathLength $MaxCombinedPdfFullPathLength",
            '"Combined PDF Name: $combinedPdfName" | Out-File $logFile -Append',
            '$pythonArgs = @($combinedPdfPath) + @($allGeneratedPdfs | ForEach-Object { [string]$_ })',
            '$mergeOutput = & $pythonExecutable $pythonScriptPath @pythonArgs 2>&1',
            '$stripOutput = & $pythonExecutable $stripPdfLayersScriptPath $combinedPdfPath 2>&1',
            '$shrinkOutput = & $pythonExecutable $shrinkScriptPath $combinedPdfPath $shrunkPath $shrinkPercentInt 2>&1',
        ):
            self.assertIn(expected, text)

        self.assertNotIn('`"$combinedPdfPath`"', text)
        self.assertNotIn('`"$shrunkPath`"', text)
        self.assertNotIn('$shrunkName = "combined-shrunk-$shrinkPercentInt-percent.pdf"', text)
        self.assertIn(
            '$shrunkName = "$combinedPdfBaseName-shrunk-$shrinkPercentInt-percent.tmp.pdf"',
            text,
        )
        self.assertIn(
            "Move-Item -LiteralPath $shrunkPath -Destination $combinedPdfPath -Force",
            text,
        )
        self.assertIn('"Shrunk PDF applied to final output: $combinedPdfName"', text)

    def test_manage_layers_script_retries_until_report_rows_verify(self):
        text = (REPO_ROOT / "scripts" / "ManageLayersDWGs.ps1").read_text(encoding="utf-8")

        for expected in (
            "function Get-ManageLayersReportRows {",
            "function Test-ManageLayersAttempt {",
            "$MaxUpdateAttemptsPerFile = 2",
            'Write-Host "PROGRESS: Processing $fileIndex of $($files.Count): $([IO.Path]::GetFileName($dwgFile)) (attempt $attempt of $MaxUpdateAttemptsPerFile)"',
            'Write-Host "PROGRESS: Verification failed for $([IO.Path]::GetFileName($dwgFile)) on attempt $attempt of ${MaxUpdateAttemptsPerFile}: $attemptReason"',
            '"DONE (verified on attempt $attempt, exit code $displayExitCode): $dwgFile" | Out-File $logFile -Append',
            '$failureSummary = "$($failedFiles.Count) of $($files.Count) file(s) failed to verify layer updates."',
            '$failureSummary = "$failureSummary First failure: $firstFailureName :: $firstFailureReason"',
            'Write-Host "PROGRESS: ERROR: $failureSummary"',
            'Write-Host "PROGRESS: Successfully updated $($files.Count) drawing(s)."',
        ):
            self.assertIn(expected, text)

    def test_manage_layers_script_uses_exact_layer_table_updates(self):
        text = (REPO_ROOT / "scripts" / "ManageLayersDWGs.ps1").read_text(encoding="utf-8")

        for expected in (
            "function ConvertTo-LispStringLiteral {",
            "$lispFreezeList = ConvertTo-LispListItems $layersToFreeze",
            "$lispThawList = ConvertTo-LispListItems $layersToThaw",
            "(defun _set-layer-frozen (layName shouldFreeze / ent rec flags nextFlags nextRec)",
            '(setq ent (tblobjname "LAYER" layName))',
            "(setq nextFlags (if shouldFreeze (logior flags 1) (- flags (logand flags 1))))",
            "(subst (cons 70 nextFlags) (assoc 70 rec) rec)",
            "(entmod nextRec)",
            'FREEZE_DXF;',
            'THAW_DXF;',
        ):
            self.assertIn(expected, text)

        self.assertNotIn('(command "_.-LAYER" "_Freeze" layName "")', text)
        self.assertNotIn('(command "_.-LAYER" "_Thaw" layName "")', text)
        self.assertNotIn('(command "_.-LAYER" "_Unlock" layName "")', text)

    def test_script_runner_preserves_specific_progress_error_on_nonzero_exit(self):
        text = MAIN_PY_PATH.read_text(encoding="utf-8")

        for expected in (
            'last_progress_error = ""',
            'if message.startswith("ERROR:"):',
            'last_progress_error = message[len(',
            'f"{last_progress_error} (exit code {return_code})."',
            'error_message = f"Script finished with error code {return_code}."',
        ):
            self.assertIn(expected, text)

    def test_script_runner_debug_output_is_safe_for_legacy_windows_codecs(self):
        text = MAIN_PY_PATH.read_text(encoding="utf-8")
        replacement_line = "AutoCAD output with replacement character: \ufffd"
        debug_line = f"DEBUG THREAD OUTPUT: {replacement_line!a}"

        self.assertIn('print(f"DEBUG THREAD OUTPUT: {line!a}")', text)
        self.assertNotIn("\ufffd", debug_line)
        debug_line.encode("cp1252")

    @unittest.skipUnless(sys.platform == "win32", "PowerShell STA relaunch is Windows-only")
    def test_sta_relaunch_preserves_files_list_path(self):
        powershell = shutil.which("powershell.exe") or shutil.which("powershell")
        self.assertIsNotNone(powershell, "powershell.exe is required to validate the STA wrapper")

        harness = textwrap.dedent(
            """
            param(
              [string]$FilesListPath = ""
            )

            if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
              $ps = (Get-Process -Id $PID).Path
              $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
              if ($PSBoundParameters.ContainsKey('FilesListPath') -and -not [string]::IsNullOrWhiteSpace($FilesListPath)) {
                $argsList += @("-FilesListPath", $FilesListPath)
              }
              $child = Start-Process -FilePath $ps -ArgumentList $argsList -Wait -PassThru
              exit $child.ExitCode
            }

            $files = @()
            $filesListWasProvided = $PSBoundParameters.ContainsKey('FilesListPath')
            $hasFilesListPath = $filesListWasProvided -and -not [string]::IsNullOrWhiteSpace($FilesListPath)
            if ($filesListWasProvided) {
              if ($hasFilesListPath) {
                Write-Host "PROGRESS: Received auto-selected files list: $FilesListPath"
                if (Test-Path $FilesListPath) {
                  $files = @(
                    Get-Content -Path $FilesListPath -Encoding UTF8 |
                      Where-Object { $_ -and $_.Trim() -and (Test-Path $_.Trim()) } |
                      ForEach-Object { $_.Trim() }
                  )
                  if ($files.Count -gt 0) {
                    Write-Host "PROGRESS: Using $($files.Count) DWG file(s) from auto-selected project folder."
                  }
                  else {
                    Write-Host "PROGRESS: Provided files list was empty. Opening file picker..."
                  }
                }
                else {
                  Write-Host "PROGRESS: Provided files list path was not found. Opening file picker..."
                }
              }
              else {
                Write-Host "PROGRESS: Files list parameter was provided without a path. Opening file picker..."
              }
            }

            if (-not $files -or $files.Count -eq 0) {
              Write-Host "PROGRESS: Waiting for user input..."
            }
            """
        ).strip()

        with tempfile.TemporaryDirectory(prefix="acies-ps-wrapper-") as temp_dir:
            temp_path = Path(temp_dir)
            dummy_dwg = temp_path / "fixture.dwg"
            dummy_dwg.write_text("", encoding="utf-8")
            files_list = temp_path / "files.txt"
            files_list.write_text(f"{dummy_dwg}\n", encoding="utf-8")
            harness_path = temp_path / "StaHarness.ps1"
            harness_path.write_text(harness, encoding="utf-8")

            result = subprocess.run(
                [
                    powershell,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(harness_path),
                    "-FilesListPath",
                    str(files_list),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
            )

            output = (result.stdout or "") + (result.stderr or "")
            self.assertEqual(0, result.returncode, msg=output)
            self.assertIn(
                f"PROGRESS: Received auto-selected files list: {files_list}",
                output,
            )
            self.assertIn(
                "PROGRESS: Using 1 DWG file(s) from auto-selected project folder.",
                output,
            )
            self.assertNotIn("PROGRESS: Waiting for user input...", output)


if __name__ == "__main__":
    unittest.main()
