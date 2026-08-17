import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class WorkItemAttachmentUiTests(unittest.TestCase):
    def test_unified_attachment_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function normalizeAttachmentEntry(entry) {",
            "function normalizeAttachments(",
            "function buildLegacyLinksFromAttachments(attachments = []) {",
            "function buildLegacyEmailRefsFromAttachments(attachments = []) {",
            "function syncAttachmentOwnerCompatFields(",
            "function getAttachmentOwnerAttachments(descriptor = {}) {",
            "async function setAttachmentOwnerAttachments(",
            "function ensureAttachmentPanel() {",
            "function openAttachmentPanel(context) {",
            "async function requestAttachmentPanelClose(options = {}) {",
            "function createAttachmentTriggerIcon(size = 15) {",
            "function createAttachmentControl(descriptor = {}, options = {}) {",
            "function createWorkItemAttachmentControls({",
            'className: "attachment-trigger",',
            'svg.setAttribute("class", "attachment-trigger-icon");',
            "trigger.appendChild(createAttachmentTriggerIcon());",
            'textContent: "Choose File",',
            'textContent: "Choose Folder",',
            'textContent: "Save Path/URL",',
            'textContent: "Choose Email File",',
            'scope: "projects-tab",',
            'scope: "edit-modal",',
        ):
            self.assertIn(expected, script)

    def test_item_level_note_and_task_attachment_payload_fields_exist(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "attachments: normalizeAttachments(task?.attachments, {",
            "attachments: normalizeAttachments(noteItem?.attachments, {",
            "task.attachments = normalizedTask.attachments;",
            "noteItem.attachments = normalizedNoteItem.attachments;",
            "attachments: normalizedTask.attachments,",
            "attachments: normalizedNoteItem.attachments,",
            "kind === \"task\"",
            "kind === \"note\"",
        ):
            self.assertIn(expected, script)

    def test_unified_attachment_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".attachment-control {",
            ".attachment-trigger {",
            ".attachment-trigger .attachment-trigger-icon {",
            ".attachment-trigger.has-attachments {",
            ".attachment-panel {",
            ".attachment-item {",
            ".attachment-type-badge {",
            ".attachment-action-row {",
            ".work-item-attachment-control {",
            ".work-item-attachment-control .attachment-trigger .attachment-trigger-icon,",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
