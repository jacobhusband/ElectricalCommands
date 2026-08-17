import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"


class ProjectDeliverableActiveUiTests(unittest.TestCase):
    def test_deliverable_active_markup_and_settings_copy_exist(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('class="d-active"', html)
        self.assertIn(">Active</label>", html)
        self.assertIn("Auto-Set Latest as Active", html)
        self.assertIn("latest due date as active", html)
        self.assertNotIn('class="d-primary"', html)
        self.assertNotIn(">Primary</label>", html)

    def test_deliverable_active_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function isDeliverableActive(deliverable) {",
            "function getProjectActiveDeliverables(project) {",
            "function getActiveAnchorDeliverable(project) {",
            "function getProjectListPriorityMeta(project) {",
            "function getProjectListPriorityDeliverable(project) {",
            "function syncProjectActiveDeliverables(project, { fallbackActiveId = \"\" } = {}) {",
            "function projectNeedsActiveMigration(sourceProject, normalizedProject) {",
            "function ensureModalProjectHasActiveDeliverable({ preferredCard = null } = {}) {",
            "active: deliverable.active === true,",
            "active: seed.active !== false,",
            "const hasExplicitActiveDeliverables = deliverables.some((deliverable) =>",
            "deliverable.active = deliverable.id === normalizedFallbackId;",
            "if (!deliverables.some((deliverable) => deliverable?.active)) {",
            "deliverables[0].active = true;",
            "const activeWithDue = activeDeliverables.filter((deliverable) =>",
            "return activeWithDue.sort(compareDeliverablesByDue)[0];",
            "const activeAnchorDeliverable = getActiveAnchorDeliverable(project);",
            "const activeIncompleteDeliverables = getProjectActiveDeliverables(project).filter(",
            "(deliverable) => !isFinished(deliverable)",
            "const activeIncompleteWithDue = activeIncompleteDeliverables.filter((deliverable) =>",
            "const priorityDeliverable = activeIncompleteWithDue.sort(compareDeliverablesByDue)[0];",
            "priorityDeliverable: activeAnchorDeliverable,",
            "hasIncompleteActiveWork: false,",
            "sortBucket: 2,",
            "fallbackDueDate: parseDueStr(getEffectiveDueStr(activeAnchorDeliverable)),",
            "hasIncompleteActiveWork: true,",
            "sortBucket: 0,",
            "sortDueDate: parseDueStr(getEffectiveDueStr(priorityDeliverable)),",
            "sortBucket: 1,",
            "return getProjectListPriorityMeta(project).priorityDeliverable;",
            'const checkedInputs = Array.from(list.querySelectorAll(".d-active:checked"));',
            'active: !!card.querySelector(".d-active")?.checked,',
            "fallbackInput.checked = true;",
            "syncProjectActiveDeliverables(project, {",
            "syncProjectActiveDeliverables(target, {",
            "syncProjectActiveDeliverables(base);",
        ):
            self.assertIn(expected, script)

        self.assertNotIn(".d-primary", script)

    def test_edit_modal_deliverables_sort_by_latest_due_date_first(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        fill_form_start = script.index("function fillForm(project) {")
        fill_form_end = script.index("function getDeliverableCardEmailRefs", fill_form_start)
        fill_form_block = script[fill_form_start:fill_form_end]

        self.assertIn(".sort(compareDeliverablesByDueDesc);", fill_form_block)
        self.assertNotIn("sortDeliverablesByPrimaryThenDueDesc(", fill_form_block)
        self.assertIn("activeAnchorDeliverable?.id", fill_form_block)


if __name__ == "__main__":
    unittest.main()
