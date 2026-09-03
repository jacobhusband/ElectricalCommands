import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import main as main_module


class XrefDwgCompareTests(unittest.TestCase):
    def test_launch_opens_new_drawing_and_runs_compare_against_archive(self):
        api = main_module.Api.__new__(main_module.Api)
        with tempfile.TemporaryDirectory(prefix="acies-xref-compare-") as temp_dir:
            root = Path(temp_dir)
            install_dir = root / "AutoCAD 2026"
            install_dir.mkdir()
            acad_exe = install_dir / "acad.exe"
            acad_core = install_dir / "accoreconsole.exe"
            acad_exe.write_bytes(b"")
            acad_core.write_bytes(b"")

            new_dwg = root / "Xrefs" / "A01-01 (E).dwg"
            old_dwg = root / "Xrefs" / "Archive" / "A01-01 (E)_old.dwg"
            new_dwg.parent.mkdir(parents=True)
            old_dwg.parent.mkdir(parents=True)
            new_dwg.write_bytes(b"new")
            old_dwg.write_bytes(b"old")

            with patch.object(
                api,
                "get_user_settings",
                return_value={"autocadPath": str(acad_core)},
            ), patch.object(
                api, "_schedule_temp_file_cleanup"
            ) as cleanup_mock, patch.object(
                main_module.subprocess,
                "Popen",
                return_value=SimpleNamespace(pid=2468),
            ) as popen_mock:
                result = api.launch_dwg_compare(str(new_dwg), str(old_dwg))

            self.assertEqual("success", result["status"])
            self.assertEqual(str(new_dwg), result["newPath"])
            self.assertEqual(str(old_dwg), result["oldPath"])
            command = popen_mock.call_args.args[0]
            self.assertEqual(str(acad_exe), command[0])
            self.assertEqual(str(new_dwg), command[1])
            self.assertEqual(["/nologo", "/b"], command[2:4])
            script_path = Path(command[4])
            self.assertEqual(str(install_dir), popen_mock.call_args.kwargs["cwd"])
            self.assertEqual(
                f'_.-COMPARE\n"{str(old_dwg).replace(chr(92), "/")}"\n',
                script_path.read_text(encoding="utf-8").replace("\r\n", "\n"),
            )
            cleanup_mock.assert_called_once_with(str(script_path))
            script_path.unlink(missing_ok=True)

    def test_launch_rejects_a_missing_archived_drawing(self):
        api = main_module.Api.__new__(main_module.Api)
        with tempfile.TemporaryDirectory(prefix="acies-xref-compare-missing-") as temp_dir:
            root = Path(temp_dir)
            new_dwg = root / "new.dwg"
            new_dwg.write_bytes(b"new")

            with patch.object(main_module.subprocess, "Popen") as popen_mock:
                result = api.launch_dwg_compare(
                    str(new_dwg),
                    str(root / "Archive" / "old.dwg"),
                )

            self.assertEqual("error", result["status"])
            self.assertIn("Archived comparison drawing was not found", result["message"])
            popen_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
