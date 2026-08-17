import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_HTML_PATH = REPO_ROOT / "index.html"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
STYLES_CSS_PATH = REPO_ROOT / "styles.css"


class HeaderAccountPopoverUiTests(unittest.TestCase):
    def test_header_account_popover_markup_exists(self):
        html = INDEX_HTML_PATH.read_text(encoding="utf-8")

        self.assertIn('id="headerAccountDropdown"', html)
        self.assertIn('id="headerGoogleAuthBtn"', html)
        self.assertIn('aria-controls="headerAccountPopover"', html)
        self.assertIn('id="headerAccountPopover"', html)
        self.assertIn('id="headerAccountPopoverAvatar"', html)
        self.assertIn('id="headerAccountPopoverName"', html)
        self.assertIn('id="headerAccountPopoverEmail"', html)
        self.assertIn('id="headerAccountPopoverStatus"', html)
        self.assertIn('id="headerAccountPopoverNote"', html)
        self.assertIn('id="headerAccountSignOutBtn"', html)
        self.assertIn('id="headerDisciplineSwitcher"', html)
        self.assertIn('role="radiogroup"', html)
        self.assertIn('data-discipline="Mechanical"', html)
        self.assertIn('data-discipline="Electrical"', html)
        self.assertIn('data-discipline="Plumbing"', html)

        dropdown_start = html.index('id="headerAccountDropdown"')
        switcher_start = html.index('id="headerDisciplineSwitcher"')
        popover_start = html.index('id="headerAccountPopover"')
        dropdown_end = html.index('</div>\n        </div>\n    </header>', dropdown_start)
        self.assertGreater(switcher_start, dropdown_start)
        self.assertLess(switcher_start, popover_start)
        self.assertLess(popover_start, dropdown_end)
        self.assertIn(
            '</button>\n                <div class="header-discipline-switcher" id="headerDisciplineSwitcher"',
            html,
        )

    def test_header_account_popover_script_wiring_exists(self):
        script = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        self.assertIn("let headerAccountPopoverOpen = false;", script)
        self.assertIn("function getGoogleAuthInitials(state = googleAuthState) {", script)
        self.assertIn("function updateHeaderAccountPopoverVisibility() {", script)
        self.assertIn("function setHeaderAccountPopoverOpen(isOpen, { focusTrigger = false } = {}) {", script)
        self.assertIn("function toggleHeaderAccountPopover() {", script)
        self.assertIn("function renderHeaderDisciplineSwitcher() {", script)
        self.assertIn("function setActiveDiscipline(discipline) {", script)
        self.assertIn("function initializeHeaderDisciplineSwitcher() {", script)
        self.assertIn("function handleHeaderGoogleAuthAction() {", script)
        self.assertIn('const headerPopoverName = document.getElementById("headerAccountPopoverName");', script)
        self.assertIn('const headerPopoverStatus = document.getElementById("headerAccountPopoverStatus");', script)
        self.assertIn('headerGoogleAuthBtn.onclick = () => handleHeaderGoogleAuthAction();', script)
        self.assertIn("initializeHeaderDisciplineSwitcher();", script)
        self.assertIn('headerAccountSignOutBtn.onclick = () => handleGoogleSignOut();', script)
        self.assertIn('setHeaderAccountPopoverOpen(false, { focusTrigger: true });', script)

    def test_header_account_popover_styles_exist(self):
        css = STYLES_CSS_PATH.read_text(encoding="utf-8")

        self.assertIn(".header-auth-dropdown {", css)
        self.assertIn("flex-direction: column;", css)
        self.assertIn("align-items: center;", css)
        self.assertIn(".header-discipline-switcher {", css)
        switcher_start = css.index(".header-discipline-switcher {")
        switcher_end = css.index("}", switcher_start)
        switcher_block = css[switcher_start:switcher_end]
        self.assertIn("margin-top: -1px;", switcher_block)
        self.assertNotIn("position: absolute;", switcher_block)
        self.assertNotIn("right: 1.5rem;", switcher_block)
        self.assertIn(".header-discipline-toggle {", css)
        self.assertIn('.header-discipline-toggle[aria-checked="true"] {', css)
        self.assertIn(".header-discipline-toggle:disabled {", css)
        self.assertIn(".header-auth-btn.is-compact {", css)
        self.assertIn("border-color: transparent;", css)
        self.assertIn(".header-auth-btn.is-compact:hover {", css)
        self.assertIn(".header-auth-btn.is-compact:focus-visible {", css)
        self.assertIn(".header-auth-btn.is-compact .header-auth-status-dot {", css)
        self.assertIn("display: none;", css)
        self.assertIn(".header-account-popover {", css)
        self.assertIn(".header-account-profile {", css)
        self.assertIn(".header-account-sync {", css)
        self.assertIn(".header-account-actions {", css)
        self.assertIn(
            '.header-account-popover[data-sync-status="active"] .header-account-sync-dot {',
            css,
        )


if __name__ == "__main__":
    unittest.main()
