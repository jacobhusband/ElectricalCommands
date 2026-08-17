import argparse
import json
import os
import shutil
import sys
import tempfile
import time
import traceback
from datetime import datetime, timezone
from pathlib import Path

import webview


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Keep backend behavior deterministic for automation runs.
os.environ.setdefault("ACIES_TEST_MODE", "1")

from main import Api, BASE_DIR, SETTINGS_FILE, TASKS_FILE  # noqa: E402


class BackupFile:
    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self._temp_copy = None
        self._did_exist = False

    def __enter__(self):
        self._did_exist = self.file_path.exists()
        if self._did_exist:
            fd, tmp_path = tempfile.mkstemp(suffix=".bak")
            os.close(fd)
            self._temp_copy = Path(tmp_path)
            shutil.copy2(self.file_path, self._temp_copy)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._did_exist and self._temp_copy and self._temp_copy.exists():
            self.file_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self._temp_copy, self.file_path)
            self._temp_copy.unlink(missing_ok=True)
        elif not self._did_exist and self.file_path.exists():
            self.file_path.unlink()
        return False


def write_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def read_json_or_default(path, default):
    path = Path(path)
    if not path.exists():
        return default
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
            return data if isinstance(data, type(default)) else default
    except Exception:
        return default


def wait_for(predicate, timeout_seconds, interval_seconds=0.2):
    start = time.time()
    last_value = None
    while time.time() - start <= timeout_seconds:
        last_value = predicate()
        if last_value:
            return last_value
        time.sleep(interval_seconds)
    raise TimeoutError(f"Timed out after {timeout_seconds}s")


def parse_js_json(window, expression):
    raw = window.evaluate_js(f"JSON.stringify(({expression}))")
    if raw in (None, "", "undefined", "null"):
        return None
    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return raw
    return raw


def build_fixture_settings(existing=None):
    settings = dict(existing or {})
    settings["discipline"] = ["Electrical"]
    settings["showSetupHelp"] = False
    settings.setdefault("apiKey", "e2e-test-key")
    settings.setdefault("autocadPath", r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe")
    return settings


def build_fixture_tasks():
    return [
        {
            "id": "DB-1001",
            "name": "Alpha Headquarters Upgrade",
            "nick": "Alpha Reno",
            "notes": "",
            "path": r"C:\Projects\AlphaHQ",
            "refs": [],
            "deliverables": [
                {
                    "id": "dlv_alpha_base",
                    "name": "IFP",
                    "due": "02/20/26",
                    "notes": "",
                    "tasks": [{"text": "Base alpha task", "done": False, "links": []}],
                    "statuses": ["Working"],
                    "statusTags": ["working"],
                    "status": "Working",
                    "emailRefs": [],
                    "emailRef": None,
                }
            ],
            "overviewDeliverableId": "dlv_alpha_base",
        },
        {
            "id": "DB-2002",
            "name": "Delta Office Relocation",
            "nick": "Delta Move",
            "notes": "",
            "path": r"C:\Projects\DeltaOffice",
            "refs": [],
            "deliverables": [
                {
                    "id": "dlv_delta_base",
                    "name": "DD50",
                    "due": "02/27/26",
                    "notes": "",
                    "tasks": [{"text": "Base delta task", "done": False, "links": []}],
                    "statuses": ["Working"],
                    "statusTags": ["working"],
                    "status": "Working",
                    "emailRefs": [],
                    "emailRef": None,
                }
            ],
            "overviewDeliverableId": "dlv_delta_base",
        },
    ]


def assert_true(condition, message):
    if not condition:
        raise AssertionError(message)


def run_automation(window, api, timeout_seconds, result_holder):
    steps = []
    try:
        wait_for(
            lambda: parse_js_json(
                window,
                "document.readyState === 'complete' && !!window.__aciesAutomation",
            ),
            timeout_seconds=timeout_seconds,
            interval_seconds=0.25,
        )
        steps.append("automation-ready")

        before_project = wait_for(
            lambda: (
                lambda project: project if project and project.get("exists") else None
            )(
                parse_js_json(
                    window, "window.__aciesAutomation.getProjectSummary(0)"
                )
            ),
            timeout_seconds=timeout_seconds,
        )
        before_count = int(before_project.get("deliverableCount", 0))
        steps.append("captured-initial-project-state")

        first_payload = {
            "id": "CLIENT-555",
            "name": "Sunset Medical Plaza Expansion",
            "due": "03/15/26",
            "path": r"C:\Client\SunsetMedical",
            "tasks": [{"text": "Update one-line diagram", "done": False}],
            "notes": "Client requested revised electrical drawings.",
        }
        first_result = parse_js_json(
            window,
            f"window.__aciesAutomation.simulateAiResult({json.dumps(first_payload)})",
        )
        assert_true(first_result and first_result.get("branch") == "no-match", "Expected no-match branch for first payload.")
        dialog_state = parse_js_json(window, "window.__aciesAutomation.getAiNoMatchDialogState()")
        assert_true(dialog_state and dialog_state.get("open"), "AI no-match dialog did not open.")
        steps.append("first-no-match-dialog-open")

        parse_js_json(window, "window.__aciesAutomation.openAiNoMatchSearch()")
        search_state = parse_js_json(
            window, "window.__aciesAutomation.setAiNoMatchSearch('alpha')"
        )
        assert_true(search_state and search_state.get("resultCount", 0) >= 1, "Expected at least one name/nickname search match.")
        select_result = parse_js_json(
            window, "window.__aciesAutomation.selectAiNoMatchProject(0)"
        )
        assert_true(
            select_result and select_result.get("selected"),
            "Failed to select existing project for AI deliverable.",
        )
        add_result = parse_js_json(
            window, "window.__aciesAutomation.confirmAiNoMatchAddToProject()"
        )
        assert_true(add_result and add_result.get("added"), "Failed to add AI deliverable to selected existing project.")
        wait_for(
            lambda: parse_js_json(window, "window.__aciesAutomation.getEditDialogState().open"),
            timeout_seconds=timeout_seconds,
        )
        edit_state_existing = parse_js_json(window, "window.__aciesAutomation.getEditDialogState()")
        assert_true(edit_state_existing and edit_state_existing.get("editIndex") == 0, "Expected edit dialog opened for selected project index 0.")
        steps.append("existing-project-add-flow")

        after_project = parse_js_json(window, "window.__aciesAutomation.getProjectSummary(0)")
        assert_true(after_project and after_project.get("exists"), "Project summary unavailable after add.")
        after_count = int(after_project.get("deliverableCount", 0))
        assert_true(after_count == before_count + 1, "Deliverable count did not increment on selected project.")
        assert_true(
            any(d.get("due") == "03/15/26" for d in after_project.get("deliverables", [])),
            "Expected deliverable due date from AI payload on selected project.",
        )
        steps.append("existing-project-deliverable-verified")

        parse_js_json(window, "window.__aciesAutomation.commitEditDialog()")
        wait_for(
            lambda: not parse_js_json(window, "window.__aciesAutomation.getEditDialogState().open"),
            timeout_seconds=timeout_seconds,
        )
        steps.append("existing-project-edit-committed")

        second_payload = {
            "id": "CLIENT-777",
            "name": "Harbor Point Utility Upgrade",
            "due": "04/01/26",
            "path": r"C:\Client\HarborPoint",
            "tasks": [{"text": "Issue updated panel schedule", "done": False}],
            "notes": "Prepare revised utility coordination package.",
        }
        second_result = parse_js_json(
            window,
            f"window.__aciesAutomation.simulateAiResult({json.dumps(second_payload)})",
        )
        assert_true(second_result and second_result.get("branch") == "no-match", "Expected no-match branch for second payload.")
        dialog_state_2 = parse_js_json(window, "window.__aciesAutomation.getAiNoMatchDialogState()")
        assert_true(dialog_state_2 and dialog_state_2.get("open"), "AI no-match dialog did not open for second payload.")
        parse_js_json(window, "window.__aciesAutomation.chooseAiNoMatchCreateNew()")
        wait_for(
            lambda: parse_js_json(window, "window.__aciesAutomation.getEditDialogState().open"),
            timeout_seconds=timeout_seconds,
        )
        edit_state_new = parse_js_json(window, "window.__aciesAutomation.getEditDialogState()")
        assert_true(
            edit_state_new and "Create Project" in (edit_state_new.get("saveButtonText") or ""),
            "Expected Create Project mode after choosing create-new path.",
        )
        assert_true(
            (edit_state_new.get("projectName") or "").strip() == second_payload["name"],
            "Expected AI project name to prefill new project form.",
        )
        steps.append("create-new-flow-opened")

        parse_js_json(window, "window.__aciesAutomation.cancelEditDialog()")
        wait_for(
            lambda: not parse_js_json(window, "window.__aciesAutomation.getEditDialogState().open"),
            timeout_seconds=timeout_seconds,
        )
        steps.append("create-new-flow-closed")

        result_holder["status"] = "success"
        result_holder["steps"] = steps
        result_holder["first_branch"] = first_result
        result_holder["second_branch"] = second_result
        result_holder["project_before"] = before_project
        result_holder["project_after"] = after_project
        result_holder["existing_edit_state"] = edit_state_existing
        result_holder["new_edit_state"] = edit_state_new
    except Exception as exc:
        result_holder["status"] = "error"
        result_holder["message"] = str(exc)
        result_holder["traceback"] = traceback.format_exc()
        result_holder["steps"] = steps
    finally:
        try:
            window.destroy()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(
        description="Run automated E2E validation for AI no-match project resolution."
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=45,
        help="Per-step timeout in seconds.",
    )
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    artifacts_dir = REPO_ROOT / "build" / "e2e" / "ai-no-match" / timestamp
    artifacts_dir.mkdir(parents=True, exist_ok=True)

    existing_settings = read_json_or_default(SETTINGS_FILE, {})
    fixture_settings = build_fixture_settings(existing_settings)
    fixture_tasks = build_fixture_tasks()
    run_result = {}

    with BackupFile(SETTINGS_FILE), BackupFile(TASKS_FILE):
        write_json(SETTINGS_FILE, fixture_settings)
        write_json(TASKS_FILE, fixture_tasks)

        api = Api()
        if not api.test_mode:
            raise RuntimeError(
                "E2E runner requires ACIES_TEST_MODE=1. Environment override failed."
            )

        window = webview.create_window(
            "ACIES E2E AI No-Match Validation",
            str(BASE_DIR / "index.html"),
            js_api=api,
            width=1400,
            height=900,
            resizable=True,
            min_size=(1024, 768),
        )
        webview.start(
            run_automation,
            args=(window, api, args.timeout, run_result),
        )

    report = {
        "status": run_result.get("status", "error"),
        "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "artifacts_dir": str(artifacts_dir),
        "result": run_result,
    }
    report_path = artifacts_dir / "report.json"
    write_json(report_path, report)

    print(f"E2E report: {report_path}")
    if report["status"] != "success":
        print(report.get("result", {}).get("message", "E2E run failed."))
        return 1

    print("E2E validation passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
