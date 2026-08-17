import re
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]


class LightingScheduleResidentialUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = (REPO_ROOT / "index.html").read_text(encoding="utf-8")
        cls.script = (REPO_ROOT / "script.js").read_text(encoding="utf-8")
        cls.styles = (REPO_ROOT / "styles.css").read_text(encoding="utf-8")

    def test_schedule_has_independent_mark_and_symbol_columns(self):
        self.assertRegex(
            self.index,
            re.compile(r"<th>MARK</th>\s*<th>SYMBOL</th>", re.MULTILINE),
        )
        self.assertIn('colspan="9"', self.index)

    def test_residential_starter_contains_all_eleven_fixture_rows(self):
        for number in range(1, 12):
            self.assertIn(f'starterFixtureKey: "ca-2025-res-l{number}"', self.script)
            self.assertTrue(
                (REPO_ROOT / "assets" / "lighting" / f"ca-residential-l{number}.png").is_file()
            )
        self.assertIn("insertCaliforniaResidentialLightingStarter", self.script)
        self.assertIn("isLightingScheduleRowBlank", self.script)

    def test_symbol_upload_and_paste_workflow_is_wired(self):
        self.assertIn("getLightingScheduleClipboardImageFiles", self.script)
        self.assertIn("saveLightingScheduleSymbol", self.script)
        self.assertIn("activateLightingScheduleSymbolPasteTarget", self.script)
        self.assertIn("handleLightingScheduleSymbolPaste", self.script)
        self.assertIn(
            'document.addEventListener("paste", handleLightingScheduleSymbolPaste)',
            self.script,
        )
        self.assertIn('td.addEventListener("click"', self.script)
        self.assertIn("window.pywebview.api.save_page_asset", self.script)
        self.assertIn("window.pywebview.api.get_page_asset", self.script)
        self.assertIn("LIGHTING_SCHEDULE_SYMBOL_MAX_BYTES", self.script)

    def test_symbol_cell_styles_exist(self):
        self.assertIn(".lighting-schedule-symbol-cell", self.styles)
        self.assertIn(".lighting-schedule-symbol-preview", self.styles)
        self.assertIn(".lighting-schedule-symbol-actions", self.styles)
        self.assertIn(".lighting-schedule-symbol-cell.is-paste-target", self.styles)

    def test_packaging_includes_bundled_lighting_symbols(self):
        root_spec = (REPO_ROOT / "ACIES Scheduler.spec").read_text(encoding="utf-8")
        build_spec = (REPO_ROOT / "build-config" / "ACIES Scheduler.spec").read_text(
            encoding="utf-8"
        )
        self.assertIn("assets\\\\lighting", root_spec)
        self.assertIn("'assets', 'lighting'", build_spec)

    def test_row_normalization_preserves_row_object_identity(self):
        # Rendered cells and the symbol paste target close over row objects;
        # normalization must mutate rows in place instead of recreating them,
        # or edits and pastes land on orphaned objects and silently vanish.
        self.assertIn(
            "function normalizeLightingScheduleRowInPlace(row)", self.script
        )
        self.assertIn("? normalizeLightingScheduleRowInPlace(row)", self.script)

    def test_symbol_rows_force_symbol_column_in_canonical_payload(self):
        self.assertIn(
            "normalized.rows.some((row) => row.symbolAssetPath || row.starterFixtureKey)",
            self.script,
        )

    def test_symbol_paste_without_target_shows_feedback(self):
        self.assertIn(
            "notifyLightingScheduleSymbolPasteWithoutTarget", self.script
        )
        self.assertIn(
            "Click a symbol cell first, then press Ctrl+V to paste the image.",
            self.script,
        )

    def test_async_sync_responses_do_not_replace_in_progress_edits(self):
        self.assertIn("let lightingScheduleEditRevision = 0;", self.script)
        self.assertIn("function markLightingScheduleDirty()", self.script)
        self.assertIn(
            "lightingScheduleEditRevision !== requestedRevision",
            self.script,
        )
        self.assertIn(
            "lightingScheduleEditRevision === savedRevision",
            self.script,
        )
        self.assertIn(
            "applyLightingScheduleRecordMetadata(schedule, response.data);",
            self.script,
        )


if __name__ == "__main__":
    unittest.main()
