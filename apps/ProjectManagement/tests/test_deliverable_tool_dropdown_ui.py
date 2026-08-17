import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class DeliverableToolDropdownUiTests(unittest.TestCase):
    def test_shared_tool_registry_has_category_field_per_entry(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "const SHARED_TOOL_LAUNCH_REGISTRY = Object.freeze([",
            'id: "toolCopyProjectLocally"',
            'menuLabel: "Work Locally"',
            'id: "toolPublishDwgs"',
            'menuLabel: "Publish"',
            'id: "toolManageLayers"',
            'menuLabel: "Freeze/Thaw"',
            'id: "toolCleanXrefs"',
            'menuLabel: "Prepare XREFs"',
            'id: "toolCreateNarrativeTemplate"',
            'menuLabel: "Create NOC"',
            'id: "toolCreatePlanCheckTemplate"',
            'menuLabel: "Create PCC"',
            'id: "toolWireSizer"',
            'id: "toolCircuitBreaker"',
            'id: "toolBackupDrawings"',
            'menuLabel: "Backup DWGs"',
            'id: "toolLightingSchedule"',
            "isReady: false,",
            'category: "general",',
            'category: "electrical",',
            'category: "templates",',
            "function getReadySharedToolLaunchEntries() {",
            "function getDeliverableToolMenuEntries() {",
            "function getDeliverableToolMenuEntriesByCategory() {",
            "const DELIVERABLE_TOOL_CATEGORIES = Object.freeze([",
            '{ key: "general", label: "General" },',
            '{ key: "electrical", label: "Electrical" },',
            '{ key: "templates", label: "Templates" },',
            "function buildProjectsTabToolLaunchContext(project, deliverable) {",
            'source: "projects-tab",',
            "function launchSharedToolCard(toolId, launchContext = null) {",
        ):
            self.assertIn(expected, text)

    def test_card_view_icon_action_wiring_exists(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function createDeliverableCardTopActions(deliverable, project, card) {",
            "function createDeliverableCardActionButton({",
            "function createDeliverableAttachmentAction(deliverable, project, card) {",
            'className: "deliverable-card-action-row"',
            'className: "deliverable-card-action-group deliverable-card-action-group-left"',
            'className: "deliverable-card-action-group deliverable-card-action-group-right"',
            "const statusDropdown = createStatusDropdown(deliverable, project, card);",
            'statusDropdown.classList.add("deliverable-card-status-action");',
            "const toolDropdown = createDeliverableToolDropdown(deliverable, project, card);",
            'toolDropdown.classList.add("deliverable-card-tool-action");',
            "const pinBtn = createDeliverablePinButton(deliverable);",
            "leftActions.append(pinBtn, statusDropdown, toolDropdown);",
            "rightActions.append(",
            "actions.append(leftActions, rightActions);",
            "openEdit(projectIndex)",
            "removeDeliverable(project, deliverable)",
            "openAttachmentPanel(attachmentContext)",
            "async function handleDeliverableActionsDrop(context, event) {",
            "await handleDeliverableActionsDrop(attachmentContext, event);",
            "updateAttachmentTriggerState(button, getAttachments());",
            "isSupportedEmailFile",
            "const actionRow = createDeliverableCardTopActions(deliverable, project, card);",
            "card.append(actionRow, header, statusSection);",
            'const actionRow = card.querySelector(":scope > .deliverable-card-action-row");',
            "card.insertBefore(projectMeta, actionRow?.nextSibling || card.firstChild);",
        ):
            self.assertIn(expected, text)

        top_actions_start = text.index("function createDeliverableCardTopActions(")
        top_actions_end = text.index("let openDeliverableActionsDropdown = null;", top_actions_start)
        top_actions_block = text[top_actions_start:top_actions_end]
        self.assertIn("leftActions.append(pinBtn, statusDropdown, toolDropdown);", top_actions_block)
        delete_index = top_actions_block.index('className: "deliverable-card-delete-action"')
        edit_index = top_actions_block.index('className: "deliverable-card-edit-action"')
        self.assertLess(delete_index, edit_index)
        # The card delete action removes the deliverable, not the whole project.
        self.assertIn('title: "Delete deliverable"', top_actions_block)
        self.assertIn("removeDeliverable(project, deliverable)", top_actions_block)
        self.assertNotIn("removeProject(projectIndex)", top_actions_block)
        # The inline attachment and open-project-page buttons have been removed.
        self.assertNotIn("attachmentBtn", top_actions_block)
        self.assertNotIn("createDeliverableAttachmentAction", top_actions_block)
        self.assertNotIn("deliverable-card-open-project-page-action", top_actions_block)
        # The project notes entry point remains.
        self.assertIn("deliverable-card-open-page-action", top_actions_block)
        self.assertNotIn("createExpandToggle(card)", top_actions_block)
        self.assertNotIn("deliverable-card-expand-action", top_actions_block)

        header_start = text.index("function createCardHeader(")
        header_end = text.index("function createExpandToggle(", header_start)
        header_block = text[header_start:header_end]
        self.assertNotIn("createDeliverableActionsDropdown", header_block)
        self.assertNotIn("actionsDropdown", header_block)
        self.assertNotIn("createDeliverableCardTopActions", header_block)

        self.assertNotIn("function createDeliverableCardBottomActions(", text)
        self.assertNotIn("const bottomActions = createDeliverableCardBottomActions(", text)
        self.assertNotIn("deliverable-card-bottom-actions", text)

    def test_card_view_icon_actions_close_other_popovers(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function closeDeliverableCardActionOverlays() {",
            "closeOpenDeliverableToolDropdown();",
            "closeOpenDeliverableStatusDropdowns();",
            "closeOpenDeliverableActionsDropdown();",
            "if (openAttachmentPanelContext) {",
            "void requestAttachmentPanelClose();",
        ):
            self.assertIn(expected, text)

    def test_card_view_icon_action_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".deliverable-card-action-row {",
            ".deliverable-card-action-group {",
            ".deliverable-card-action-group-left {",
            ".deliverable-card-action-group-right {",
            ".deliverable-card-action-btn,",
            ".deliverable-card-action-btn:hover,",
            ".deliverable-card-action-btn.is-pinned {",
            ".deliverable-card-action-btn.danger {",
            ".deliverable-card-action-btn.is-dragover {",
            ".deliverable-card-action-btn.has-attachments::after {",
            ".deliverable-card-action-row .deliverable-status-trigger,",
            ".deliverable-card-action-row .deliverable-tool-menu {",
        ):
            self.assertIn(expected, css)

        self.assertNotIn(".deliverable-card-bottom-actions {", css)
        self.assertNotIn(".deliverable-card-top-actions {", css)


if __name__ == "__main__":
    unittest.main()
