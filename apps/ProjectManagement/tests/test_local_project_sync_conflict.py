import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import json
import tempfile
import shutil
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

# Stubs for GUI libraries in main.py
import sys
import types

def _ensure_stubs():
    # google.genai
    try:
        from google import genai as _genai
        from google.genai import types as _types
    except Exception:
        google_module = types.ModuleType("google")
        google_module.__path__ = []
        sys.modules["google"] = google_module
        genai_module = types.ModuleType("google.genai")
        genai_types_module = types.ModuleType("google.genai.types")
        genai_module.types = genai_types_module
        google_module.genai = genai_module
        sys.modules["google.genai"] = genai_module
        sys.modules["google.genai.types"] = genai_types_module

    # webview
    try:
        import webview
    except Exception:
        webview_module = types.ModuleType("webview")
        webview_module.windows = []
        webview_module.create_window = lambda *args, **kwargs: None
        webview_module.start = lambda *args, **kwargs: None
        sys.modules["webview"] = webview_module

    # dotenv
    try:
        from dotenv import load_dotenv
    except Exception:
        dotenv_module = types.ModuleType("dotenv")
        dotenv_module.load_dotenv = lambda *args, **kwargs: False
        sys.modules["dotenv"] = dotenv_module

    # requests
    try:
        import requests
    except Exception:
        requests_module = types.ModuleType("requests")
        requests_module.get = lambda *args, **kwargs: None
        requests_module.post = lambda *args, **kwargs: None
        sys.modules["requests"] = requests_module

    # pydantic
    try:
        from pydantic import BaseModel, Field
    except Exception:
        pydantic_module = types.ModuleType("pydantic")
        class BaseModel: pass
        def Field(*args, **kwargs):
            return args[0] if args else kwargs.get("default")
        pydantic_module.BaseModel = BaseModel
        pydantic_module.Field = Field
        sys.modules["pydantic"] = pydantic_module

_ensure_stubs()

import main
from main import Api

class LocalProjectSyncConflictTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.local_dir = os.path.join(self.temp_dir, "local")
        self.server_dir = os.path.join(self.temp_dir, "server")
        os.makedirs(self.local_dir)
        os.makedirs(self.server_dir)
        
        self.temp_metadata_file = os.path.join(self.temp_dir, "sync_metadata.json")
        self.patcher = patch("main.SYNC_METADATA_FILE", self.temp_metadata_file)
        self.patcher.start()
        
        self.api = Api.__new__(Api)
        self.api.get_user_settings = MagicMock(return_value={})
        
    def tearDown(self):
        self.patcher.stop()
        shutil.rmtree(self.temp_dir)

    def test_sync_initial_files_and_detect_conflict(self):
        # Create a file on server
        test_file = "Electrical/test.txt"
        local_test_file = os.path.join(self.local_dir, test_file)
        server_test_file = os.path.join(self.server_dir, test_file)
        os.makedirs(os.path.dirname(server_test_file), exist_ok=True)
        
        with open(server_test_file, "w") as f:
            f.write("server version")
            
        # Copy to local
        res = self.api._apply_local_project_manager_direction(
            self.local_dir, self.server_dir, [test_file], direction="to_local"
        )
        self.assertEqual(res["status"], "success")
        self.assertTrue(os.path.exists(local_test_file))
        
        # Verify metadata is written
        self.assertTrue(os.path.exists(self.temp_metadata_file))
        
        # Modify both local and server to introduce conflict
        future_time = os.path.getmtime(local_test_file) + 10.0
        with open(local_test_file, "w") as f:
            f.write("local modification")
        os.utime(local_test_file, (future_time, future_time))

        with open(server_test_file, "w") as f:
            f.write("server modification")
        os.utime(server_test_file, (future_time, future_time))
            
        # Run compare
        comparison = self.api._compare_local_project_manager_files(self.local_dir, self.server_dir)
        conflict_candidates = comparison.get("conflictCandidates", [])
        self.assertEqual(len(conflict_candidates), 1)
        self.assertEqual(os.path.normpath(conflict_candidates[0]["relativePath"]), os.path.normpath(test_file))
        
        # Resolve conflict - keeping local
        res_resolve = self.api.resolve_local_project_manager_conflict(
            self.local_dir, self.server_dir, test_file, "keep_local"
        )
        self.assertEqual(res_resolve["status"], "success")
        
        # Verify server is updated to local version
        with open(server_test_file, "r") as f:
            self.assertEqual(f.read(), "local modification")
            
        # Verification: run compare again. Conflict should be gone and files should be equal.
        comparison2 = self.api._compare_local_project_manager_files(self.local_dir, self.server_dir)
        self.assertEqual(len(comparison2.get("conflictCandidates", [])), 0)
        self.assertEqual(len(comparison2.get("equalFiles", [])), 1)

    def test_sync_deletion_propagation(self):
        # Create a file, sync to local
        test_file = "Electrical/delete_me.txt"
        local_test_file = os.path.join(self.local_dir, test_file)
        server_test_file = os.path.join(self.server_dir, test_file)
        os.makedirs(os.path.dirname(server_test_file), exist_ok=True)
        
        with open(server_test_file, "w") as f:
            f.write("content")
            
        self.api._apply_local_project_manager_direction(
            self.local_dir, self.server_dir, [test_file], direction="to_local"
        )
        
        # Delete local file
        os.remove(local_test_file)
        
        # Compare project files
        comparison = self.api._compare_local_project_manager_files(self.local_dir, self.server_dir)
        self.assertEqual(comparison["status"], "success")
        
        # We expect a delete candidate
        local_to_server = comparison.get("localToServerCandidates", [])
        delete_candidates = [x for x in local_to_server if x.get("reason") == "local_deleted"]
        self.assertEqual(len(delete_candidates), 1)
        self.assertEqual(os.path.normpath(delete_candidates[0]["relativePath"]), os.path.normpath(test_file))
        
        # Sync deletion to server
        res_sync = self.api._apply_local_project_manager_direction(
            self.local_dir, self.server_dir, [test_file], direction="to_server"
        )
        self.assertEqual(res_sync["status"], "success")
        
        # Verify server file is deleted
        self.assertFalse(os.path.exists(server_test_file))

class LocalProjectManagerComparisonScopeTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.local_dir = os.path.join(self.temp_dir, "local")
        self.server_dir = os.path.join(self.temp_dir, "server")
        os.makedirs(self.local_dir)
        os.makedirs(self.server_dir)

        self.api = Api.__new__(Api)
        self.api.get_user_settings = MagicMock(return_value={})

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def _write(self, root, rel_path, content, mtime=None):
        full_path = os.path.join(root, rel_path)
        os.makedirs(os.path.dirname(full_path), exist_ok=True)
        with open(full_path, "w") as f:
            f.write(content)
        if mtime is not None:
            os.utime(full_path, (mtime, mtime))
        return full_path

    def test_newer_file_outside_managed_folder_is_detected(self):
        # "Reports" is not a discipline folder or Xrefs -> additive_only scope.
        rel_path = "Reports/notes.txt"
        server_file = self._write(self.server_dir, rel_path, "server version")
        base_mtime = os.path.getmtime(server_file)
        # Make the local copy clearly newer than the server copy.
        self._write(self.local_dir, rel_path, "local version", mtime=base_mtime + 600.0)
        os.utime(server_file, (base_mtime, base_mtime))

        comparison = self.api._compare_local_project_manager_files(
            self.local_dir, self.server_dir
        )
        self.assertEqual(comparison["status"], "success")
        newer = [
            item
            for item in comparison.get("localToServerCandidates", [])
            if str(item.get("changeType") or "").lower() == "newer"
        ]
        self.assertEqual(len(newer), 1)
        self.assertEqual(
            os.path.normpath(newer[0]["relativePath"]), os.path.normpath(rel_path)
        )
        self.assertEqual(newer[0]["scopeType"], "additive_only")
        self.assertTrue(newer[0]["selectedByDefault"])


class CopyProjectSourcePathResolutionTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.api = Api.__new__(Api)
        self.api.get_user_settings = MagicMock(return_value={})

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_workroom_uses_saved_path_without_six_digit_id(self):
        # A real project folder that does not follow the 6-digit-ID naming convention.
        project_dir = os.path.join(self.temp_dir, "Maple Street Renovation")
        os.makedirs(project_dir)
        launch_context = {
            "source": "workroom",
            "projectPath": project_dir,
            "rootProjectPath": project_dir,
        }

        resolution = self.api._resolve_copy_project_source_path(
            None, launch_context, {}
        )
        self.assertEqual(resolution["status"], "success")
        self.assertEqual(
            os.path.normpath(resolution["path"]), os.path.normpath(project_dir)
        )


if __name__ == "__main__":
    unittest.main()
