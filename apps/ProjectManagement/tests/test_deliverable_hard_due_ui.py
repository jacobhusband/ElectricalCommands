import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
INDEX_HTML_PATH = REPO_ROOT / "index.html"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class DeliverableHardDueUiTests(unittest.TestCase):
    @staticmethod
    def _block(text, start_marker, end_marker):
        start = text.index(start_marker)
        end = text.index(end_marker, start)
        return text[start:end]

    def test_data_model_carries_hard_due(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        normalize_block = self._block(
            script,
            "function normalizeDeliverable(",
            "function createDeliverable(",
        )
        self.assertIn('due: String(deliverable.due || "").trim(),', normalize_block)
        self.assertIn(
            'hardDue: String(deliverable.hardDue || "").trim(),', normalize_block
        )

        create_block = self._block(
            script,
            "function createDeliverable(",
            "function normalizeProject(",
        )
        self.assertIn('hardDue: seed.hardDue || "",', create_block)

    def test_due_state_helpers_expose_four_states(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function getEffectiveDueStr(deliverable) {",
            "function getHardDueStr(deliverable) {",
            "function deliverableDueState(deliverable) {",
            "function isDeliverableHardDueMissed(deliverable) {",
        ):
            self.assertIn(expected, script)

        state_block = self._block(
            script,
            "function deliverableDueState(",
            "function isDeliverableHardDueMissed(",
        )
        # A missed hard deadline escalates above plain overdue, but never for
        # work that is already finished.
        self.assertIn('return "critical";', state_block)
        self.assertIn("!isFinished(deliverable)", state_block)
        self.assertIn("return dueState(getEffectiveDueStr(deliverable));", state_block)

        effective_block = self._block(
            script,
            "function getEffectiveDueStr(",
            "function getHardDueStr(",
        )
        # Internal date wins, hard date is the fallback.
        self.assertIn('String(deliverable?.due || "").trim()', effective_block)
        self.assertIn('String(deliverable?.hardDue || "").trim()', effective_block)

    def test_scheduling_logic_uses_effective_due_date(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            # Sorting
            "function compareDeliverablesByDue(a, b) {\n  const da = parseDueStr(getEffectiveDueStr(a));",
            # Timeframe filter
            "function matchesDueFilter(deliverable, filter) {\n  if (filter === \"all\") return true;\n  const d = parseDueStr(getEffectiveDueStr(deliverable));",
            # Week / kanban view
            "function deliverableDueInWeek(deliverable, weekStart) {\n  const d = parseDueStr(getEffectiveDueStr(deliverable));",
            "function deliverableIsOverdueIncomplete(deliverable, weekStart) {\n  const d = parseDueStr(getEffectiveDueStr(deliverable));",
            # Project rows
            "dueDate: parseDueStr(getEffectiveDueStr(deliverable)),",
            "isHardDueMissed: isDeliverableHardDueMissed(deliverable),",
        ):
            self.assertIn(expected, script)

    def test_missed_hard_deadline_sorts_above_plain_overdue(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        compare_block = self._block(
            script,
            "function compareProjectDeliverableRows(",
            "function sortProjectDeliverableRows(",
        )
        self.assertIn("const aHardMissed = !!a?.isHardDueMissed;", compare_block)
        self.assertIn("const bHardMissed = !!b?.isHardDueMissed;", compare_block)
        self.assertIn(
            "if (aHardMissed !== bHardMissed) return aHardMissed ? -1 : 1;",
            compare_block,
        )

    def test_pin_urgent_deliverables_catches_hard_deadlines(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        pin_block = self._block(
            script,
            "async function pinUrgentDeliverables(",
            "function resetCopyProjectLocallyDialogState(",
        )
        self.assertIn("const hardDueStr = getHardDueStr(deliverable);", pin_block)
        self.assertIn(
            'const hardUrgent = !!parseDueStr(hardDueStr) && dueState(hardDueStr) !== "ok";',
            pin_block,
        )
        self.assertIn("if (!hardUrgent) {", pin_block)

    def test_deliverable_card_renders_a_separate_hard_badge(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("function createDeliverableDueBadges(deliverable, project) {", script)
        badges_block = self._block(
            script,
            "function createDeliverableDueBadges(",
            "function createExpandToggle(",
        )
        self.assertIn('field: "due",', badges_block)
        self.assertIn('field: "hardDue",', badges_block)
        self.assertIn('stateClass: missed ? "hard critical" : "hard",', badges_block)
        self.assertIn("const hardDue = getHardDueStr(deliverable);", badges_block)
        self.assertIn("Past hard deadline", badges_block)

    def test_badge_calendar_writes_to_the_field_it_was_opened_from(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        calendar_block = self._block(
            script,
            "function showCalendarForDeliverableBadge(",
            "function renderInlineCalendar(",
        )
        self.assertIn('field = "due"', calendar_block)
        self.assertIn("deliverable[field] = formatDueDateShort(selectedDate);", calendar_block)
        self.assertNotIn("deliverable.due = formatDueDateShort(", calendar_block)

    def test_edit_modal_exposes_a_hard_deadline_input(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('<label class="label">Internal Due Date</label>', html)
        self.assertIn('<label class="label">Hard Deadline</label>', html)
        self.assertIn('<input class="d-hard-due" placeholder="MM/DD/YYYY" />', html)
        self.assertIn("Must-finish date. Leave blank if this can be pushed.", html)

    def test_modal_populates_validates_and_saves_hard_due(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn(
            'card.querySelector(".d-hard-due").value = deliverable.hardDue || "";',
            script,
        )
        self.assertIn(
            "const dateInput = wrapper.querySelector('.d-due, .d-hard-due');", script
        )
        self.assertIn(
            "const inputs = document.querySelectorAll('.d-due, .d-hard-due');", script
        )
        self.assertIn(
            "document.querySelector('.d-due.input-error, .d-hard-due.input-error')",
            script,
        )

        read_form_block = self._block(
            script,
            "function readForm(",
            "function addRefRowFrom(",
        )
        self.assertIn(
            'const hardDue = card.querySelector(".d-hard-due").value.trim();',
            read_form_block,
        )
        self.assertIn("hardDue,", read_form_block)

    def test_internal_date_after_hard_deadline_warns_without_blocking(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        order_block = self._block(
            script,
            "function validateDeliverableDateOrder(",
            "function validateAllDueDates(",
        )
        self.assertIn("if (!due || !hardDue || due <= hardDue) return true;", order_block)
        self.assertIn("dueInput.classList.add('input-warning');", order_block)
        self.assertIn(
            "'Warning: Internal date is after the hard deadline'", order_block
        )
        # Warning only - never blocks the save.
        self.assertNotIn("return false;", order_block)

    def test_hard_and_critical_badge_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".deliverable-due-badge.hard {", css)
        self.assertIn(".deliverable-due-badge.hard.critical,", css)
        self.assertIn(".deliverable-due-badge.critical {", css)
        self.assertIn(".deliverable-summary-due {", css)

    def test_note_level_due_date_ui_stays_removed(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertNotIn("function createNoteDueDateControl(", script)
        self.assertNotIn(".note-due-badge {", css)


if __name__ == "__main__":
    unittest.main()
