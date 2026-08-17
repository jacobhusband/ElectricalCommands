import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class PanelScheduleFilenameUiTests(unittest.TestCase):
    def setUp(self):
        self.text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

    def test_launch_context_reader_exists(self):
        self.assertIn("function getLaunchContextProjectId(launchContext = null) {", self.text)
        self.assertIn('return String(launchContext.projectId || "").trim();', self.text)

    def test_launch_contexts_carry_project_id(self):
        # The two CAD-tool launch-context builders both expose projectId.
        for marker in (
            "function buildWorkroomCadLaunchContext() {",
            "function buildProjectsTabToolLaunchContext(project, deliverable) {",
        ):
            start = self.text.index(marker)
            body = self.text[start : self.text.index("\n}", start)]
            self.assertIn(
                'projectId: String(project?.id || "").trim(),',
                body,
                f"{marker} should carry projectId",
            )

    def test_new_schedule_default_name_uses_project_id(self):
        self.assertIn(
            "const projectId = getLaunchContextProjectId(circuitBreakerState.launchContext);",
            self.text,
        )
        self.assertIn(
            'const safeId = projectId.replace(/[\\\\/:*?"<>|]/g, "").trim();',
            self.text,
        )
        self.assertIn(
            'const baseName = safeId ? `${safeId} Panel Schedules` : "Panel_Schedule";',
            self.text,
        )
        self.assertIn("        baseName,\n        outputExtension", self.text)


if __name__ == "__main__":
    unittest.main()
