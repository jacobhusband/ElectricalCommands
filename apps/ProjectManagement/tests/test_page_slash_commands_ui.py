import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class PageSlashCommandsUiTests(unittest.TestCase):
    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_visible_toolbar_removed_and_slash_menu_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn('id="pageSlashMenu"', html)
        self.assertIn('role="listbox"', html)
        self.assertNotIn('class="page-toolbar"', html)
        self.assertNotIn("data-page-cmd=", html)
        self.assertNotIn("data-page-block=", html)
        self.assertNotIn('id="pageTextColorInput"', html)
        self.assertNotIn('id="pageInsertImageBtn"', html)

        self.assertIn(".page-slash-menu {", css)
        self.assertIn(".page-slash-menu-header {", css)
        self.assertIn(".page-slash-menu-item {", css)
        self.assertIn(".page-slash-menu-item.is-active {", css)
        self.assertIn(".page-slash-menu-shortcut {", css)
        self.assertIn(".page-slash-color-swatch {", css)
        self.assertNotIn(".page-toolbar {", css)
        self.assertNotIn(".page-fmt-btn {", css)

    def test_command_registry_contains_block_inline_media_and_project_page_commands(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        registry = self._block(
            script,
            "const PAGE_SLASH_COMMANDS = Object.freeze([",
            "const PAGE_IMAGE_MIN_PERCENT = 10;",
        )

        self.assertIn('id: "text"', registry)
        self.assertIn('action: { type: "block", tag: "p" }', registry)
        self.assertIn('id: "h1"', registry)
        self.assertIn('action: { type: "block", tag: "h1" }', registry)
        self.assertIn('id: "h2"', registry)
        self.assertIn('id: "h3"', registry)
        self.assertIn('id: "bullet"', registry)
        self.assertIn('command: "insertUnorderedList"', registry)
        self.assertIn('id: "numbered"', registry)
        self.assertIn('command: "insertOrderedList"', registry)
        self.assertIn('id: "todo"', registry)
        self.assertIn('aliases: ["todo", "to-do", "checkbox", "checklist", "task"]', registry)
        self.assertIn('action: { type: "todo" }', registry)
        self.assertIn('id: "bold"', registry)
        self.assertIn('command: "bold"', registry)
        self.assertIn('id: "italic"', registry)
        self.assertIn('aliases: ["italic", "italics", "i"]', registry)
        self.assertIn('id: "underline"', registry)
        self.assertIn('id: "link"', registry)
        self.assertIn('action: { type: "link" }', registry)
        self.assertIn('id: "color"', registry)
        self.assertIn('action: { type: "color" }', registry)
        self.assertIn('id: "image"', registry)
        self.assertIn('action: { type: "image" }', registry)
        self.assertIn('id: "page"', registry)
        self.assertIn("projectOnly: true", registry)

    def test_slash_menu_filters_project_only_commands_and_handles_keyboard(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        filter_block = self._block(
            script,
            "function getPageSlashCommands(query = \"\") {",
            "function positionPageSlashMenu(menu, range) {",
        )
        range_block = self._block(
            script,
            "function getPageSlashRangeFromCaret(editor",
            "function getPageSlashCommands(query = \"\") {",
        )
        keydown_block = self._block(
            script,
            "function handlePageSlashKeydown(event) {",
            "// --- Image insertion / hydration",
        )
        update_block = self._block(
            script,
            "function updatePageSlashMenuFromCaret() {",
            "function getPageSlashCommandShortcut(command) {",
        )

        self.assertIn("allowInlineTrigger = false", range_block)
        self.assertIn("&& !allowInlineTrigger", range_block)
        self.assertIn("getRangeTextBeforeCaretInSlashBlock(range, editor)", range_block)
        self.assertIn("function getPageSlashBlockForRange(range, editor) {", script)
        self.assertIn("node.matches?.(\"p,div,h1,h2,h3,li,blockquote,pre\")", script)
        self.assertIn("if (command.projectOnly && !pageNav.project) return false;", filter_block)
        self.assertIn("command.aliases.some((alias)", filter_block)
        self.assertNotIn("!pageSlashState.open) return", update_block)
        self.assertIn("if (!pageSlashState.open) pageSlashState.preservedSelection = null;", update_block)
        self.assertIn("pageSlashState.open = true;", update_block)
        self.assertIn("pageSlashState.query = slash.query;", update_block)
        self.assertIn('event.key === "Escape"', keydown_block)
        self.assertIn("closePageSlashMenu();", keydown_block)
        self.assertIn('event.key === "ArrowDown" || event.key === "ArrowUp"', keydown_block)
        self.assertIn('event.key === "Enter" || event.key === "Tab"', keydown_block)
        self.assertIn("executePageSlashCommand(commands[pageSlashState.selectedIndex]);", keydown_block)

    def test_slash_input_preserves_selected_text_and_deletes_query_before_execution(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        beforeinput_block = self._block(
            script,
            "function handlePageSlashBeforeInput(event) {",
            "function handlePageSlashKeydown(event) {",
        )
        execute_block = self._block(
            script,
            "function executePageSlashCommand(command, options = {}) {",
            "function createProjectSubpageFromSlash() {",
        )

        self.assertIn('event.inputType !== "insertText" || event.data !== "/"', beforeinput_block)
        self.assertIn("getRangeTextBeforeCaretInSlashBlock(range, editor)", beforeinput_block)
        self.assertIn("if (!range.collapsed) {", beforeinput_block)
        self.assertIn("event.preventDefault();", beforeinput_block)
        self.assertIn("const preservedSelection = clonePageRange(range);", beforeinput_block)
        self.assertIn('const slashNode = document.createTextNode("/");', beforeinput_block)
        self.assertIn("slashRange.insertNode(slashNode);", beforeinput_block)
        self.assertIn("selection.addRange(caretRange);", beforeinput_block)
        self.assertIn("openPageSlashMenu({ range: queryRange, preservedSelection });", beforeinput_block)
        self.assertIn("deletePageSlashQuery();", execute_block)
        self.assertIn("closePageSlashMenu();", execute_block)
        self.assertIn("selection.addRange(preservedSelection);", execute_block)
        self.assertIn("allowInlineTrigger: pageSlashState.open && !!pageSlashState.preservedSelection", script)

    def test_slash_query_deletion_anchors_caret_in_emptied_block(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        delete_block = self._block(
            script,
            "function deletePageSlashQuery() {",
            "function executePageSlashCommand(command, options = {}) {",
        )

        # The target line/block must be captured before the query text is removed.
        self.assertIn("const block = getPageSlashBlockForRange(range, editor);", delete_block)
        self.assertIn("range.deleteContents();", delete_block)
        # An emptied block makes the caret resolve into the previous block, so a
        # <br> placeholder keeps block/list commands on the current line.
        self.assertIn(
            'if (block && block !== editor && !block.textContent && !block.querySelector("br, img")) {',
            delete_block,
        )
        self.assertIn('block.appendChild(document.createElement("br"));', delete_block)
        self.assertIn("range.setStart(block, 0);", delete_block)

    def test_command_preview_shows_header_query_and_shortcuts(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        render_block = self._block(
            script,
            "function renderPageSlashMenu() {",
            "function renderPageSlashColorRow(menu) {",
        )

        self.assertIn('className: "page-slash-menu-header"', render_block)
        self.assertIn('textContent: "Commands"', render_block)
        self.assertIn('textContent: pageSlashState.query ? `/${pageSlashState.query}` : "Type a command"', render_block)
        self.assertIn("function getPageSlashCommandShortcut(command) {", script)
        self.assertIn('className: "page-slash-menu-shortcut"', render_block)
        self.assertIn("textContent: getPageSlashCommandShortcut(command)", render_block)

    def test_commands_route_to_existing_editor_helpers(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        execute_block = self._block(
            script,
            "function executePageSlashCommand(command, options = {}) {",
            "function createProjectSubpageFromSlash() {",
        )

        self.assertIn("applyPageBlock(command.action.tag);", execute_block)
        self.assertIn("applyPageCommand(command.action.command);", execute_block)
        self.assertIn("insertPageTodo();", execute_block)
        self.assertIn("insertPageLink();", execute_block)
        self.assertIn("applyPageTextColor(options.color || PAGE_SLASH_COLORS[0]);", execute_block)
        self.assertIn("openPageImagePicker();", execute_block)
        self.assertIn("createProjectSubpageFromSlash();", execute_block)
        self.assertIn("function resetPageBlockAfterHeading(event) {", script)
        self.assertIn('document.execCommand("formatBlock", false, "p");', script)

    def test_todo_command_creates_persisted_checklist_items(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")
        todo_block = self._block(
            script,
            "function createPageChecklistItem({ checked = false, text = \"\" } = {}) {",
            "// --- Slash commands ---",
        )
        serialize_block = self._block(
            script,
            "function serializePageHtml(editor) {",
            "// Route page persistence",
        )

        self.assertIn('item.className = "page-checklist-item";', todo_block)
        self.assertIn('checkbox.className = "page-checklist-checkbox";', todo_block)
        self.assertIn('textEl.className = "page-checklist-text";', todo_block)
        self.assertIn("function syncPageChecklistItem(item) {", todo_block)
        self.assertIn('checkbox.setAttribute("checked", "checked");', todo_block)
        self.assertIn('item.classList.add("is-complete");', todo_block)
        self.assertIn("function isPageChecklistTextEmpty(textEl) {", todo_block)
        self.assertIn("function isPageSelectionAtChecklistTextEnd(textEl) {", todo_block)
        self.assertIn("function replacePageChecklistItemWithParagraph(item) {", todo_block)
        self.assertIn("function insertPageTodo() {", todo_block)
        self.assertIn("function handlePageChecklistKeydown(event) {", todo_block)
        self.assertIn("if (isPageChecklistTextEmpty(textEl)) {", todo_block)
        self.assertIn("replacePageChecklistItemWithParagraph(item);", todo_block)
        self.assertIn("if (!isPageSelectionAtChecklistTextEnd(textEl)) return false;", todo_block)
        self.assertIn("syncPageChecklistItems(editor);", serialize_block)
        self.assertIn(".page-checklist-item {", css)
        self.assertIn(".page-checklist-checkbox {", css)
        self.assertIn(".page-checklist-item.is-complete .page-checklist-text {", css)

    def test_page_lists_indent_with_tab_after_slash_menu_key_handling(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        list_block = self._block(
            script,
            "function handlePageListIndentKeydown(event) {",
            "// --- Slash commands ---",
        )
        ready_block = self._block(
            script,
            "function ensurePageViewReady() {",
            "init();",
        )

        self.assertIn('event.key !== "Tab"', list_block)
        self.assertIn("event.ctrlKey || event.metaKey || event.altKey", list_block)
        self.assertIn("const range = getPageSelectionRange(editor);", list_block)
        self.assertIn('element?.closest?.("li")', list_block)
        self.assertIn('item?.closest?.("ul,ol")', list_block)
        self.assertIn('document.execCommand(event.shiftKey ? "outdent" : "indent", false, null);', list_block)
        self.assertIn("savePageSelection(editor);", list_block)
        self.assertIn("queuePageSave();", list_block)
        self.assertIn("if (handlePageSlashKeydown(e)) return;", ready_block)
        self.assertIn("if (handlePageListIndentKeydown(e)) return;", ready_block)
        self.assertLess(
            ready_block.index("if (handlePageSlashKeydown(e)) return;"),
            ready_block.index("if (handlePageListIndentKeydown(e)) return;"),
        )

    def test_color_menu_and_project_page_command_create_child_subpage(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        color_block = self._block(
            script,
            "function renderPageSlashColorRow(menu) {",
            "function deletePageSlashQuery() {",
        )
        page_block = self._block(
            script,
            "function createProjectSubpageFromSlash() {",
            "function handlePageSlashBeforeInput(event) {",
        )

        self.assertIn("PAGE_SLASH_COLORS.forEach((color) => {", color_block)
        self.assertIn('className: "page-slash-color-swatch"', color_block)
        self.assertIn('type: "color"', color_block)
        self.assertIn('color: custom.value', color_block)

        self.assertIn("const { project, subpage } = pageNav;", page_block)
        self.assertIn("if (!project) return;", page_block)
        self.assertIn("pageEditorTarget.html = serializePageHtml(editor);", page_block)
        self.assertIn("pageEditorTarget.updatedAt = new Date().toISOString();", page_block)
        self.assertIn('title: "Untitled"', page_block)
        self.assertIn("parentId: subpage?.id || null", page_block)
        self.assertIn("subpages.push(child);", page_block)
        self.assertNotIn("projectSubpageExpandedIds", page_block)
        self.assertIn("openProjectPage(project, child);", page_block)

    def test_page_view_wires_slash_events_to_page_editor_only(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        ready_block = self._block(
            script,
            "function ensurePageViewReady() {",
            "init();",
        )

        self.assertIn("const editor = getPageEditorEl();", ready_block)
        self.assertIn('editor.addEventListener("beforeinput", handlePageSlashBeforeInput);', ready_block)
        self.assertIn("updatePageSlashMenuFromCaret();", ready_block)
        self.assertIn('editor.addEventListener("change", (e) => {', ready_block)
        self.assertIn("syncPageChecklistItem(checkbox.closest(\".page-checklist-item\"));", ready_block)
        self.assertIn("if (handlePageSlashKeydown(e)) return;", ready_block)
        self.assertIn("if (handlePageListIndentKeydown(e)) return;", ready_block)
        self.assertIn("if (handlePageChecklistKeydown(e)) return;", ready_block)
        self.assertIn("resetPageBlockAfterHeading(e);", ready_block)
        self.assertNotIn("pageTitle.addEventListener", ready_block)


if __name__ == "__main__":
    unittest.main()
