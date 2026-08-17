import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"
INDEX_HTML_PATH = REPO_ROOT / "index.html"


class ExpenseAttachmentUiTests(unittest.TestCase):
    def test_custom_expense_ui_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="addCustomExpenseBtn"', html)
        self.assertIn('<dialog id="customExpenseDlg">', html)
        self.assertIn('id="custom_expense_project_name"', html)
        self.assertIn('id="custom_expense_job_number"', html)
        self.assertIn('id="btnSaveCustomExpense"', html)
        self.assertIn("function openAddCustomExpenseDialog() {", script)
        self.assertIn("function saveCustomExpense() {", script)
        self.assertIn('source: "custom"', script)
        self.assertIn('textContent: getExpenseProjectSecondaryLabel(project)', script)
        self.assertIn('return isCustomExpenseProject(project) ? "Custom expense" : "Job #: --";', script)
        self.assertIn('openAddCustomExpenseDialog();', script)
        self.assertIn('saveCustomExpense();', script)

    def test_expense_attachment_preview_script_wiring_exists(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("const EXPENSE_IMAGE_THUMB_MAX_SIZE = 320;", text)
        self.assertIn("const EXPENSE_IMAGE_MODAL_MAX_SIZE = 1800;", text)
        self.assertIn("window.pywebview.api.get_expense_image_preview", text)
        self.assertIn("window.pywebview.api.resolve_expense_attachment_path", text)
        self.assertIn("function hydrateExpenseImageThumb(previewButton, image) {", text)
        self.assertIn("function openExpenseImagePreview(image) {", text)
        self.assertIn("function openImagePreviewDialog({ dataUrl, filename, width, height }) {", text)
        self.assertIn("function closeImagePreviewDialog() {", text)
        self.assertIn("function getImagePreviewStageSize(imageWidth, imageHeight) {", text)
        self.assertIn("function resetImagePreviewTransform() {", text)
        self.assertIn("function applyImagePreviewZoom(nextScale, clientX, clientY) {", text)
        self.assertIn("function applyImagePreviewPan(deltaX, deltaY) {", text)
        self.assertIn("stage.style.width = `${stageSize.width}px`;", text)
        self.assertIn("stage.style.removeProperty(\"width\");", text)
        self.assertIn("function openExpenseAttachment(image) {", text)
        self.assertIn('event.preventDefault();', text)
        self.assertIn('event.stopPropagation();', text)
        self.assertNotIn("function setExpenseImagePreviewDialogState(", text)
        self.assertNotIn("expenseImagePreviewTitle", text)
        self.assertNotIn("expenseImagePreviewOpenBtn", text)
        self.assertNotIn("expenseImagePreviewStatus", text)

    def test_expense_attachment_preview_dialog_markup_exists(self):
        text = INDEX_HTML_PATH.read_text(encoding="utf-8")
        self.assertIn('<dialog id="imagePreviewDlg"', text)
        self.assertIn('id="imagePreviewCloseBtn"', text)
        self.assertIn('id="imagePreviewStage"', text)
        self.assertIn('id="imagePreviewImg"', text)
        self.assertNotIn('id="imagePreviewTitle"', text)
        self.assertNotIn('id="imagePreviewOpenBtn"', text)
        self.assertNotIn('id="imagePreviewStatus"', text)

    def test_expense_attachment_preview_styles_exist(self):
        text = STYLES_CSS_PATH.read_text(encoding="utf-8")
        self.assertIn(".expense-image-preview-btn {", text)
        self.assertIn(".expense-attachment-placeholder {", text)
        self.assertIn(".image-preview-dialog {", text)
        self.assertIn("position: fixed;", text)
        self.assertIn("inset: 50% auto auto 50%;", text)
        self.assertIn("transform: translate(-50%, -50%);", text)
        self.assertIn("margin: 0;", text)
        self.assertIn(".image-preview-close {", text)
        self.assertIn("transform: translate(38%, -38%);", text)
        self.assertIn(".image-preview-stage {", text)
        self.assertNotIn("width: min(96vw, 1400px);", text)
        self.assertIn("max-width: calc(100vw - 2rem);", text)
        self.assertIn(".image-preview-stage.is-zoomed {", text)
        self.assertIn(".image-preview-stage.is-dragging {", text)
        self.assertIn("@keyframes expense-thumb-sheen {", text)


if __name__ == "__main__":
    unittest.main()
