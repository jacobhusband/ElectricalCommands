import shutil
import subprocess
import sys
import tempfile
import unittest
import json
import zipfile
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = REPO_ROOT / "scripts" / "removeXREFPaths.ps1"


@unittest.skipUnless(sys.platform == "win32", "removeXREFPaths.ps1 behavior tests are Windows-only")
class RemoveXrefPathsBehaviorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.powershell = shutil.which("powershell.exe") or shutil.which("powershell")
        if cls.powershell is None:
            raise unittest.SkipTest("powershell.exe is required to validate removeXREFPaths.ps1")

    def _write_file(self, path, content):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")

    def _run_script(self, selected_paths, discipline_short="E"):
        files_list_path = selected_paths[0].parent / "__selected_files.txt"
        files_list_path.write_text(
            "\n".join(str(path) for path in selected_paths),
            encoding="utf-8",
        )

        result = subprocess.run(
            [
                self.powershell,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(SCRIPT_PATH),
                "-AcadCore",
                self.powershell,
                "-DisciplineShort",
                discipline_short,
                "-FilesListPath",
                str(files_list_path),
                "-SkipAcad",
                "1",
                "-StripXrefs",
                "0",
                "-SetByLayer",
                "0",
                "-Purge",
                "0",
                "-Audit",
                "0",
                "-HatchColor",
                "0",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
        )
        output = (result.stdout or "") + (result.stderr or "")
        return result, output

    def _run_script_with_manifest(self, manifest_path, discipline_short="E"):
        result = subprocess.run(
            [
                self.powershell,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(SCRIPT_PATH),
                "-AcadCore",
                self.powershell,
                "-DisciplineShort",
                discipline_short,
                "-SourceManifestPath",
                str(manifest_path),
                "-SkipAcad",
                "1",
                "-StripXrefs",
                "0",
                "-SetByLayer",
                "0",
                "-Purge",
                "0",
                "-Audit",
                "0",
                "-HatchColor",
                "0",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60,
        )
        output = (result.stdout or "") + (result.stderr or "")
        return result, output

    def test_arch_source_stages_into_xrefs_and_archives_existing_target(self):
        with tempfile.TemporaryDirectory(prefix="acies-remove-xrefs-arch-") as temp_dir:
            project_root = Path(temp_dir) / "260243 Example"
            arch_source = project_root / "Arch" / "A01-01 plan.dwg"
            canonical_target = project_root / "Xrefs" / "A01-01 (E).dwg"
            self._write_file(arch_source, "arch-source")
            self._write_file(canonical_target, "existing-target")

            result, output = self._run_script([arch_source])

            self.assertEqual(0, result.returncode, msg=output)
            self.assertTrue(arch_source.exists())
            self.assertEqual("arch-source", canonical_target.read_text(encoding="utf-8"))
            self.assertFalse(list((project_root / "Xrefs").glob("__incoming__*.dwg")))
            archived_targets = list((project_root / "Xrefs" / "Archive").glob("A01-01 (E)_*.dwg"))
            self.assertEqual(1, len(archived_targets), msg=output)
            self.assertEqual("existing-target", archived_targets[0].read_text(encoding="utf-8"))
            self.assertIn("Staged Arch source in Xrefs as __incoming__", output)
            self.assertIn("Archived existing file to A01-01 (E)_", output)

    def test_manifest_file_source_stages_into_xrefs(self):
        with tempfile.TemporaryDirectory(prefix="acies-remove-xrefs-manifest-file-") as temp_dir:
            project_root = Path(temp_dir) / "260247 Example"
            arch_source = project_root / "Arch" / "A05-05 plan.dwg"
            canonical_target = project_root / "Xrefs" / "A05-05 (E).dwg"
            self._write_file(arch_source, "manifest-arch-source")
            manifest_path = project_root / "manifest.json"
            manifest_path.write_text(
                json.dumps([{"kind": "file", "path": str(arch_source)}]),
                encoding="utf-8",
            )

            result, output = self._run_script_with_manifest(manifest_path)

            self.assertEqual(0, result.returncode, msg=output)
            self.assertEqual("manifest-arch-source", canonical_target.read_text(encoding="utf-8"))
            self.assertIn("Using 1 DWG source(s) from workflow selection.", output)
            self.assertIn("Staged Arch source in Xrefs as __incoming__", output)

    def test_manifest_zip_entry_extracts_and_stages_into_project_xrefs(self):
        with tempfile.TemporaryDirectory(prefix="acies-remove-xrefs-manifest-zip-") as temp_dir:
            project_root = Path(temp_dir) / "260248 Example"
            arch_folder = project_root / "Arch" / "Latest"
            arch_folder.mkdir(parents=True, exist_ok=True)
            zip_path = arch_folder / "CAD.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("Backgrounds/A06-06 plan.dwg", "zip-arch-source")
            canonical_target = project_root / "Xrefs" / "A06-06 (E).dwg"
            manifest_path = project_root / "manifest.json"
            manifest_path.write_text(
                json.dumps([
                    {
                        "kind": "zipEntry",
                        "zipPath": str(zip_path),
                        "entryName": "Backgrounds/A06-06 plan.dwg",
                        "projectRoot": str(project_root),
                        "displayPath": f"{zip_path}::Backgrounds/A06-06 plan.dwg",
                    }
                ]),
                encoding="utf-8",
            )

            result, output = self._run_script_with_manifest(manifest_path)

            self.assertEqual(0, result.returncode, msg=output)
            self.assertEqual("zip-arch-source", canonical_target.read_text(encoding="utf-8"))
            self.assertFalse(list((project_root / "Xrefs").glob("__incoming__*.dwg")))
            self.assertIn("Preparing Arch ZIP source for Xrefs transfer: A06-06 plan.dwg", output)
            self.assertIn("Staged Arch ZIP source in Xrefs as __incoming__", output)

    def test_xrefs_source_is_archived_before_canonical_target_is_created(self):
        with tempfile.TemporaryDirectory(prefix="acies-remove-xrefs-source-") as temp_dir:
            project_root = Path(temp_dir) / "260244 Example"
            xrefs_source = project_root / "Xrefs" / "A02-02 background.dwg"
            canonical_target = project_root / "Xrefs" / "A02-02 (E).dwg"
            self._write_file(xrefs_source, "xrefs-source")

            result, output = self._run_script([xrefs_source])

            self.assertEqual(0, result.returncode, msg=output)
            self.assertFalse(xrefs_source.exists())
            self.assertTrue(canonical_target.exists())
            self.assertEqual("xrefs-source", canonical_target.read_text(encoding="utf-8"))
            self.assertFalse(list((project_root / "Xrefs").glob("__incoming__*.dwg")))
            archived_sources = list((project_root / "Xrefs" / "Archive").glob("A02-02 background_*.dwg"))
            self.assertEqual(1, len(archived_sources), msg=output)
            self.assertEqual("xrefs-source", archived_sources[0].read_text(encoding="utf-8"))
            self.assertIn("Staged Xrefs source in Xrefs as __incoming__", output)
            self.assertIn("Archived selected Xrefs source to A02-02 background_", output)

    def test_xrefs_source_already_named_as_canonical_target_is_recreated_after_backup(self):
        with tempfile.TemporaryDirectory(prefix="acies-remove-xrefs-canonical-") as temp_dir:
            project_root = Path(temp_dir) / "260245 Example"
            canonical_target = project_root / "Xrefs" / "A03-03 (E).dwg"
            self._write_file(canonical_target, "current-canonical")

            result, output = self._run_script([canonical_target])

            self.assertEqual(0, result.returncode, msg=output)
            self.assertTrue(canonical_target.exists())
            self.assertEqual("current-canonical", canonical_target.read_text(encoding="utf-8"))
            self.assertFalse(list((project_root / "Xrefs").glob("__incoming__*.dwg")))
            archived_sources = list((project_root / "Xrefs" / "Archive").glob("A03-03 (E)_*.dwg"))
            self.assertEqual(1, len(archived_sources), msg=output)
            self.assertEqual("current-canonical", archived_sources[0].read_text(encoding="utf-8"))
            self.assertIn("Archived selected Xrefs source to A03-03 (E)_", output)

    def test_xrefs_archive_selection_is_rejected_without_creating_a_new_target(self):
        with tempfile.TemporaryDirectory(prefix="acies-remove-xrefs-archive-") as temp_dir:
            project_root = Path(temp_dir) / "260246 Example"
            archived_source = project_root / "Xrefs" / "Archive" / "A04-04 old.dwg"
            canonical_target = project_root / "Xrefs" / "A04-04 (E).dwg"
            self._write_file(archived_source, "archived-copy")

            result, output = self._run_script([archived_source])

            self.assertEqual(0, result.returncode, msg=output)
            self.assertTrue(archived_source.exists())
            self.assertEqual("archived-copy", archived_source.read_text(encoding="utf-8"))
            self.assertFalse(canonical_target.exists())
            self.assertIn(
                "Selected DWG is already inside the Xrefs\\Archive folder.",
                output,
            )
            self.assertIn("ERROR: 1 of 1 file(s) failed to process.", output)

    def test_cleanup_lisp_recolors_hatches_via_chprop_not_entmod(self):
        text = SCRIPT_PATH.read_text(encoding="utf-8")

        # Hatch recolor must go through CHPROP; hand-editing HATCH entity data
        # with entmod produced malformed entities that triggered AutoCAD
        # recovery prompts when the saved drawings were reopened.
        self.assertIn('(command "_.CHPROP" target "" "_Color" "9" "")', text)
        self.assertNotIn("(entmod entData)", text)
        self.assertNotIn("(cons 62 9)", text)
        # Space switching uses setvar instead of layout-switch commands.
        self.assertIn('(setvar "CTAB" "Model")', text)
        self.assertNotIn('(command "_.MODEL")', text)

    def test_script_verifies_processed_drawings_reopen_cleanly(self):
        text = SCRIPT_PATH.read_text(encoding="utf-8")

        self.assertIn("function Invoke-AcadCoreScript {", text)
        self.assertIn("function Get-AuditErrorCount {", text)
        self.assertIn("function Get-AcadOutputFailureSignal {", text)
        self.assertIn("verify_AUDIT.scr", text)
        self.assertIn("Verified clean on reopen:", text)
        self.assertIn("AUDIT error(s) on reopen.", text)

    def test_script_uses_direct_dwg_only_manual_selection(self):
        text = SCRIPT_PATH.read_text(encoding="utf-8")

        self.assertIn('function Show-DwgFileDialog {', text)
        self.assertIn('[string]$SourceManifestPath = ""', text)
        self.assertIn('function Read-SourceManifest {', text)
        self.assertIn('function Resolve-SourceItemToWorkingSource {', text)
        self.assertIn('Write-Host "PROGRESS: Opening DWG file picker..."', text)
        self.assertIn('$dlg.Filter = "DWG files (*.dwg)|*.dwg"', text)
        self.assertNotIn("Select ZIP or DWG source file(s)", text)
        self.assertNotIn("ZIP source selected. Extracting archive", text)
        self.assertNotIn('"{0}_Prepared" -f $zipBaseName', text)


if __name__ == "__main__":
    unittest.main()
