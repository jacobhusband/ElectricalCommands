import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class PageCanvasUiTests(unittest.TestCase):
    """Canvas pages support multi-select, group drag, and self-sizing text boxes."""

    @staticmethod
    def _block(text: str, start_marker: str, end_marker: str) -> str:
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_text_nodes_persist_their_auto_width_flag(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        normalize_block = self._block(
            script,
            "function normalizeCanvasNode(raw, usedIds) {",
            "function normalizeCanvasEdge(",
        )
        self.assertIn("node.autoW = raw.autoW !== false;", normalize_block)

        create_block = self._block(
            script,
            "function createCanvasTextNode(worldPt, { edit = true } = {}) {",
            "function deleteCanvasNode(nodeId) {",
        )
        self.assertIn("autoW: true,", create_block)
        self.assertIn("setCanvasNodeSelection([node.id]);", create_block)

    def test_auto_width_nodes_skip_the_inline_width(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function isCanvasNodeAutoWidth(node) {", script)
        self.assertIn('return node.type === "text" && node.autoW !== false;', script)

        apply_block = self._block(
            script,
            "function applyCanvasNodeWidth(node, el) {",
            "function renderCanvasNode(node) {",
        )
        self.assertIn('el.classList.add("is-auto-width");', apply_block)
        self.assertIn('el.style.width = "";', apply_block)
        self.assertIn('el.classList.remove("is-auto-width");', apply_block)
        self.assertIn("el.style.width = `${node.w}px`;", apply_block)

        render_block = self._block(
            script,
            "function renderCanvasNode(node) {",
            "function renderCanvasNodes() {",
        )
        self.assertIn("applyCanvasNodeWidth(node, el);", render_block)
        self.assertIn("if (isCanvasNodeSelected(node.id)) {", render_block)

        styles = STYLES_CSS_PATH.read_text(encoding="utf-8")
        auto_width_block = self._block(
            styles,
            ".page-canvas-node.is-auto-width {",
            ".page-canvas-node-text {",
        )
        self.assertIn("width: max-content;", auto_width_block)
        self.assertIn("max-width: 420px;", auto_width_block)

    def test_resize_handle_pins_width_and_double_click_restores_auto(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        move_block = self._block(
            script,
            "function handleCanvasPointerMove(e) {",
            "function handleCanvasPointerUp(e) {",
        )
        self.assertIn('if (node.type === "text") {', move_block)
        self.assertIn("node.autoW = false;", move_block)

        dblclick_block = self._block(
            script,
            "function handleCanvasDblClick(e) {",
            "function handleCanvasWheel(e) {",
        )
        self.assertIn('if (e.target.closest(".page-canvas-node-handle")) {', dblclick_block)
        self.assertIn("node.autoW = true;", dblclick_block)
        self.assertIn("applyCanvasNodeWidth(node, nodeEl);", dblclick_block)

    def test_text_editing_uses_plain_text_and_keeps_enter_local(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        editable_block = self._block(
            script,
            "function setCanvasTextEditable(textEl, editable) {",
            "function enterCanvasNodeEdit(nodeId) {",
        )
        self.assertIn('textEl.contentEditable = "plaintext-only";', editable_block)
        self.assertIn('textEl.contentEditable = "true";', editable_block)

        enter_block = self._block(
            script,
            "function enterCanvasNodeEdit(nodeId) {",
            "function createCanvasTextNode(",
        )
        self.assertIn("if (pageCanvasState.editingNodeId === nodeId) return;", enter_block)
        self.assertIn("setCanvasTextEditable(textEl, true);", enter_block)

        keydown_block = self._block(
            script,
            "function handleCanvasKeyDown(e) {",
            "function handleCanvasDblClick(e) {",
        )
        self.assertIn('if (e.key === "Enter") e.stopPropagation();', keydown_block)

        commit_block = self._block(
            script,
            "function commitCanvasNodeEdit() {",
            "function setCanvasTextEditable(",
        )
        self.assertIn("if (node && el && isCanvasNodeAutoWidth(node)) {", commit_block)
        self.assertIn("if (measured > 0) node.w = measured;", commit_block)

    def test_selection_state_holds_multiple_nodes_and_edges(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("selection: { nodes: [], edges: [] },", script)
        self.assertIn("function getCanvasSelection() {", script)
        self.assertIn("function getSelectedCanvasNodeIds() {", script)
        self.assertIn("function getSelectedCanvasNodes() {", script)
        self.assertIn("function isCanvasNodeSelected(nodeId) {", script)
        self.assertIn("function isCanvasEdgeSelected(edgeId) {", script)
        self.assertIn("function hasCanvasSelection() {", script)
        self.assertIn("function refreshCanvasSelectionClasses() {", script)
        self.assertIn("function setCanvasSelection({ nodes = [], edges = [] } = {}) {", script)
        self.assertIn("function setCanvasNodeSelection(nodeIds) {", script)
        self.assertIn("function clearCanvasSelection() {", script)
        self.assertIn("function toggleCanvasNodeSelection(nodeId) {", script)
        self.assertIn("function toggleCanvasEdgeSelection(edgeId) {", script)
        # The old single-selection shape is gone.
        self.assertNotIn('pageCanvasState.selection?.type === "node"', script)
        self.assertNotIn('setCanvasSelection({ type: "node", id:', script)
        self.assertNotIn("setCanvasSelection(null)", script)

    def test_delete_removes_the_whole_selection_in_one_pass(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function deleteCanvasItems(nodeIds, edgeIds) {", script)
        self.assertIn("function deleteCanvasSelection() {", script)

        items_block = self._block(
            script,
            "function deleteCanvasItems(nodeIds, edgeIds) {",
            "function deleteCanvasSelection() {",
        )
        self.assertIn("const doomedNodes = new Set(nodeIds);", items_block)
        self.assertIn("const doomedEdges = new Set(edgeIds);", items_block)
        self.assertIn(
            "!doomedEdges.has(edge.id) && !doomedNodes.has(edge.from) && !doomedNodes.has(edge.to)",
            items_block,
        )

        keydown_block = self._block(
            script,
            "function handleCanvasKeyDown(e) {",
            "function handleCanvasDblClick(e) {",
        )
        self.assertIn("deleteCanvasSelection();", keydown_block)
        self.assertIn("setCanvasNodeSelection(canvas.nodes.map((node) => node.id));", keydown_block)

    def test_empty_space_drag_marquee_selects_instead_of_panning(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function getPageCanvasMarqueeEl() {", script)
        self.assertIn("function updateCanvasMarqueeOverlay(gesture, e) {", script)
        self.assertIn("function applyCanvasMarqueeSelection(gesture, e) {", script)
        self.assertIn("function endCanvasMarquee() {", script)

        marquee_block = self._block(
            script,
            "function applyCanvasMarqueeSelection(gesture, e) {",
            "function endCanvasMarquee() {",
        )
        self.assertIn(
            "rect.x < maxX && rect.x + rect.w > minX && rect.y < maxY && rect.y + rect.h > minY",
            marquee_block,
        )
        self.assertIn("nodes: gesture.additive ? [...gesture.baseNodes, ...hits] : hits,", marquee_block)

        down_block = self._block(
            script,
            "function handleCanvasPointerDown(e) {",
            "function handleCanvasPointerMove(e) {",
        )
        self.assertIn('mode: "marquee",', down_block)
        self.assertIn("startWorld: getCanvasWorldPoint(e),", down_block)
        self.assertIn("if (!additive) clearCanvasSelection();", down_block)
        # Panning only happens on middle-drag or space-drag.
        self.assertIn("const panRequested = e.button === 1 || pageCanvasState.spaceDown;", down_block)
        self.assertIn("if (panRequested) {", down_block)

        move_block = self._block(
            script,
            "function handleCanvasPointerMove(e) {",
            "function handleCanvasPointerUp(e) {",
        )
        self.assertIn('if (gesture.mode === "marquee") {', move_block)
        self.assertIn("applyCanvasMarqueeSelection(gesture, e);", move_block)

        up_block = self._block(
            script,
            "function handleCanvasPointerUp(e) {",
            "function cancelCanvasGesture() {",
        )
        self.assertIn('if (gesture.mode === "marquee") {', up_block)
        self.assertIn("endCanvasMarquee();", up_block)

        cancel_block = self._block(
            script,
            "function cancelCanvasGesture() {",
            "function handleCanvasKeyDown(e) {",
        )
        self.assertIn(
            "setCanvasSelection({ nodes: gesture.baseNodes, edges: gesture.baseEdges });",
            cancel_block,
        )

    def test_ctrl_click_toggles_and_dragging_moves_every_selected_node(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("function isCanvasAdditiveEvent(e) {", script)
        self.assertIn("return !!(e.ctrlKey || e.metaKey || e.shiftKey);", script)

        down_block = self._block(
            script,
            "function handleCanvasPointerDown(e) {",
            "function handleCanvasPointerMove(e) {",
        )
        self.assertIn("const additive = isCanvasAdditiveEvent(e);", down_block)
        self.assertIn("if (!panRequested && node && additive) {", down_block)
        self.assertIn("toggleCanvasNodeSelection(node.id);", down_block)
        self.assertIn("const wasSelected = isCanvasNodeSelected(node.id);", down_block)
        self.assertIn(
            "const wasSoleSelection = wasSelected && getSelectedCanvasNodeIds().length === 1;",
            down_block,
        )
        self.assertIn("const dragging = getSelectedCanvasNodes();", down_block)
        self.assertIn(
            "nodes: dragging.map((item) => ({ id: item.id, origX: item.x, origY: item.y })),",
            down_block,
        )

        move_block = self._block(
            script,
            "function handleCanvasPointerMove(e) {",
            "function handleCanvasPointerUp(e) {",
        )
        self.assertIn("gesture.nodes.forEach((entry) => {", move_block)
        self.assertIn("node.x = entry.origX + dx;", move_block)
        self.assertIn("node.y = entry.origY + dy;", move_block)

        up_block = self._block(
            script,
            "function handleCanvasPointerUp(e) {",
            "function cancelCanvasGesture() {",
        )
        self.assertIn("if (gesture.nodes.length > 1) {", up_block)
        self.assertIn("setCanvasNodeSelection([gesture.clickedId]);", up_block)
        self.assertIn(
            'else if (gesture.wasSoleSelection && clicked?.type === "text") {',
            up_block,
        )
        self.assertIn("enterCanvasNodeEdit(gesture.clickedId);", up_block)

    def test_marquee_overlay_markup_and_styles_exist(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")
        self.assertIn(
            '<div class="page-canvas-marquee" id="pageCanvasMarquee" hidden></div>',
            html,
        )
        self.assertIn("Drag to select, Ctrl+click to add to the selection", html)

        styles = STYLES_CSS_PATH.read_text(encoding="utf-8")
        marquee_block = self._block(
            styles,
            ".page-canvas-marquee {",
            ".page-canvas-viewport.is-drop-target::after {",
        )
        self.assertIn("position: absolute;", marquee_block)
        self.assertIn("pointer-events: none;", marquee_block)
        self.assertIn(".page-canvas-marquee[hidden] {", marquee_block)

    def test_canvas_state_reset_clears_the_marquee(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        reset_block = self._block(
            script,
            "function resetPageCanvasState() {",
            "function queuePageCanvasSave() {",
        )
        self.assertIn("selection: { nodes: [], edges: [] },", reset_block)
        self.assertIn("const marquee = getPageCanvasMarqueeEl();", reset_block)
        self.assertIn("if (marquee) marquee.hidden = true;", reset_block)
        self.assertIn("hideCanvasContextMenu();", reset_block)

    def test_canvas_context_menu_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "let canvasContextMenu = null;",
            "let canvasContextMenuHandlersReady = false;",
            "function hideCanvasContextMenu() {",
            "function ensureCanvasContextMenu() {",
            "function showCanvasContextMenu(event) {",
            'textContent: "Convert to Panel Schedule with AI",',
            'data-canvas-menu-action": "panelSchedule",',
            "function getSelectedCanvasImageNodes() {",
            "function getSelectedCanvasTextNodes() {",
            "void openCanvasPanelScheduleDialog();",
        ):
            self.assertIn(expected, script)

    def test_context_menu_item_is_disabled_without_image_nodes(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        show_block = self._block(
            script,
            "function showCanvasContextMenu(event) {",
            "function handleCanvasContextMenu(e) {",
        )
        self.assertIn("const imageCount = getSelectedCanvasImageNodes().length;", show_block)
        self.assertIn("item.disabled = imageCount === 0 && !hasPendingWork;", show_block)
        # Reachable with nothing selected while an analysis runs or findings wait.
        self.assertIn("!!canvasPanelScheduleState.job || !!canvasPanelScheduleState.pendingResult;", show_block)
        self.assertIn('"Review panel schedule findings"', show_block)
        self.assertIn(
            'menu.querySelector("button[data-canvas-menu-action]:not([disabled])")?.focus();',
            show_block,
        )

    def test_right_click_selects_the_node_under_the_cursor(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        menu_block = self._block(
            script,
            "function handleCanvasContextMenu(e) {",
            "function handleCanvasPointerDown(e) {",
        )
        self.assertIn(
            "if (nodeId && !isCanvasNodeSelected(nodeId)) setCanvasNodeSelection([nodeId]);",
            menu_block,
        )
        self.assertIn("commitCanvasNodeEdit();", menu_block)
        self.assertIn("showCanvasContextMenu(e);", menu_block)
        # The box being edited keeps its native menu (spell-check, paste).
        self.assertIn('e.target?.closest?.(".page-canvas-node-text")', menu_block)

    def test_canvas_registers_the_contextmenu_listener(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        ready_block = self._block(
            script,
            "function ensurePageCanvasReady() {",
            "function renderPageCanvasView() {",
        )
        self.assertIn(
            'viewport.addEventListener("contextmenu", handleCanvasContextMenu);',
            ready_block,
        )

    def test_context_menu_dismiss_uses_pointerdown_capture(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        ensure_block = self._block(
            script,
            "function ensureCanvasContextMenu() {",
            "function showCanvasContextMenu(event) {",
        )
        self.assertIn('document.addEventListener(\n      "pointerdown",', ensure_block)
        self.assertIn('document.addEventListener("scroll", hideCanvasContextMenu, true);', ensure_block)

    def test_canvas_context_menu_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        for expected in (
            ".canvas-context-menu {",
            "position: fixed;",
            ".canvas-context-menu[hidden] {",
            ".canvas-context-menu__item {",
            ".canvas-context-menu__item:disabled {",
        ):
            self.assertIn(expected, css)


if __name__ == "__main__":
    unittest.main()
