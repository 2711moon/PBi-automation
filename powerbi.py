"""
powerbi.py -- Headless Power BI automation via Playwright.

Responsibilities:
  1. Log into Power BI with provided credentials (handles MFA OTP via terminal)
  2. Navigate to the configured workspace and report (no hardcoded URLs)
  3. For each AOM, apply a URL filter and export all pages as PDF

Runs headless (invisible) - no browser window appears during execution.
"""
import os
import re
import logging
from typing import Optional

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from config import EXPORTS_DIR

log = logging.getLogger(__name__)

# Timeouts (milliseconds)
NAV_TIMEOUT    = 60_000
REPORT_WAIT_MS = 12_000   # extra wait for all visuals to render before export
EXPORT_TIMEOUT = 180_000  # PDF generation can be slow on large reports


class PowerBIExporter:
    """
    Context manager: launches headless browser, logs in, discovers report URL,
    then exports a filtered PDF for each AOM.
    """

    def __init__(self, username: str, password: str):
        self.username     = username
        self.password     = password
        self._pw          = None
        self._browser     = None
        self._context     = None
        self._page        = None
        self._report_urls  = {}   # cache: (workspace, report_name) -> url
        self._slicer_cache    = {}   # cache: report_name -> slicer index that contains AOM names
        self._aom_trigger_idx = None  # cached index of the AOM slicer in the trigger list

    # -- Context manager -------------------------------------------------------

    def __enter__(self) -> "PowerBIExporter":
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-gpu",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ]
        )
        self._context = self._browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/127.0.0.0 Safari/537.36"
            ),
        )
        self._page = self._context.new_page()

        log.info("Browser started (headless). Logging into Power BI...")
        self._login()
        # Report URL discovery is now done dynamically per report in export_report.
        return self

    def __exit__(self, *args) -> None:
        try:
            if self._browser:
                self._browser.close()
            if self._pw:
                self._pw.stop()
        except Exception:
            pass

    # -- Login -----------------------------------------------------------------

    def _login(self) -> None:
        """
        Authenticate with Microsoft login.
        Handles MFA by prompting for OTP or push-notification approval in terminal.
        """
        page = self._page

        # Go directly to Microsoft login â€” skips the app.powerbi.com SSO redirect chain
        log.info("  Navigating to Microsoft login page...")
        page.goto("https://login.microsoftonline.com/", timeout=NAV_TIMEOUT, wait_until="domcontentloaded")

        page.wait_for_timeout(1_500)

        # Step 1: Enter email
        # Microsoft uses multiple possible selectors depending on tenant config
        EMAIL_SEL = 'input[name="loginfmt"], input[type="email"], #i0116'
        try:
            page.wait_for_selector(EMAIL_SEL, timeout=NAV_TIMEOUT)
        except PWTimeout:
            os.makedirs(EXPORTS_DIR, exist_ok=True)
            debug = os.path.join(EXPORTS_DIR, "debug_login_page.png")
            page.screenshot(path=debug)
            raise RuntimeError(
                f"Login page did not load as expected. Current URL: {page.url}\n"
                f"Screenshot saved to: {debug}"
            )

        page.fill(EMAIL_SEL, self.username)
        page.keyboard.press("Enter")
        page.wait_for_timeout(2_000)
        log.info("  Email submitted.")

        # Step 2: Enter password
        PASS_SEL = 'input[name="passwd"], input[type="password"], #i0118'
        page.wait_for_selector(PASS_SEL, timeout=NAV_TIMEOUT)
        page.fill(PASS_SEL, self.password)
        page.keyboard.press("Enter")
        page.wait_for_timeout(2_000)
        log.info("  Password submitted. Checking for MFA...")

        # Step 3: Handle MFA (if enabled)
        self._handle_mfa()

        # Step 4: "Stay signed in?" -- always click No for automation
        try:
            page.wait_for_selector("#idBtn_Back", timeout=8_000)
            page.click("#idBtn_Back")
            log.info("  'Stay signed in?' -- answered No.")
        except PWTimeout:
            pass

        # Step 5: Now that login is done, navigate explicitly to Power BI
        # (Microsoft may redirect to M365 home â€” this ensures we land on Power BI)
        log.info("  Navigating to Power BI...")
        page.goto("https://app.powerbi.com/", timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(3_000)

        # Step 6: Handle potential Power BI sign-up / confirmation interstitial
        # Sometimes Power BI asks to re-enter email: "Enter your work or school email..."
        try:
            interstitial_email = page.locator('input[placeholder="Enter email"], input[type="email"]').first
            if interstitial_email.is_visible(timeout=3_000):
                log.info("  Power BI interstitial detected. Re-submitting email...")
                interstitial_email.fill(self.username)
                
                # Try clicking Submit button if it exists
                submit_btn = page.locator('button:has-text("Submit")').first
                if submit_btn.is_visible(timeout=1_000):
                    submit_btn.click()
                else:
                    page.keyboard.press("Enter")
                    
                page.wait_for_timeout(5_000)
        except Exception:
            pass

        if "app.powerbi.com" not in page.url:
            os.makedirs(EXPORTS_DIR, exist_ok=True)
            debug = os.path.join(EXPORTS_DIR, "debug_post_login.png")
            page.screenshot(path=debug)
            raise RuntimeError(
                f"Could not reach Power BI after login. URL: {page.url}\n"
                f"Screenshot saved to: {debug}"
            )

        log.info("  Logged in successfully.")

    def _handle_mfa(self) -> None:
        """
        Detect which MFA challenge Microsoft presented and handle it:
          - TOTP / SMS code  -> prompt user to type the code in terminal
          - Push notification -> prompt user to approve on phone, then press Enter
        """
        page = self._page
        page.wait_for_timeout(2_500)   # give Microsoft a moment to show MFA screen

        # Check for a code-input field (TOTP or SMS)
        try:
            code_input = page.locator('input[name="otc"], input[autocomplete="one-time-code"]').first
            if code_input.is_visible(timeout=5_000):
                print()
                print("  MFA required -- enter the 6-digit code from your Authenticator app or SMS:")
                otp = input("  OTP Code : ").strip()
                code_input.fill(otp)
                page.keyboard.press("Enter")
                log.info("  OTP entered.")
                page.wait_for_timeout(3_000)
                return
        except PWTimeout:
            pass

        # Check for push-notification screen
        try:
            push_text = page.locator(
                'text=Open your Microsoft Authenticator, '
                'text=Approve the sign-in, '
                'text=approve sign-in request'
            )
            if push_text.first.is_visible(timeout=5_000):
                print()
                print("  MFA required -- approve the sign-in request on your Authenticator app.")
                input("  Press Enter here after approving on your phone... ")
                page.wait_for_timeout(3_000)
                log.info("  Push notification approved by user.")
                return
        except PWTimeout:
            pass

        # No MFA screen detected -- continue normally
        log.info("  No MFA challenge detected.")

    # -- Identity verification prompt ------------------------------------------

    def _handle_identity_prompt(self) -> None:
        """
        Power BI shows 'You need to verify your identity' dialog at various points.
        Always click 'Continue'. Loops up to 5 times to catch re-appearances.
        """
        for attempt in range(1, 6):
            try:
                btn = self._page.locator(
                    'button:has-text("Continue"), '
                    '[aria-label="Continue"]'
                ).first
                btn.wait_for(timeout=3_000, state="visible")
                btn.click()
                log.info(f"  Identity prompt dismissed (attempt {attempt}).")
                self._page.wait_for_timeout(1_500)
            except PWTimeout:
                break

    def _dismiss_popups(self) -> None:
        """
        Dismiss any overlay dialogs that Power BI or Copilot shows:
          - Copilot onboarding dialog (has an X close button)
          - Any other CDK overlay modal
        Presses Escape multiple times as a general catch-all.
        """
        page = self._page

        # Close Copilot onboarding/welcome dialog if visible
        try:
            close_btn = page.locator(
                'button[aria-label="Close"], '
                'button[aria-label="Dismiss"], '
                'button[data-testid="dismiss-button"]'
            ).first
            close_btn.wait_for(timeout=3_000, state="visible")
            close_btn.click(force=True)
            log.info("  Dismissed popup dialog.")
            page.wait_for_timeout(800)
        except PWTimeout:
            pass

        # Escape up to 3 times to clear any CDK overlay backdrops
        for _ in range(3):
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)


    # -- Report URL Discovery --------------------------------------------------


    def relogin(self) -> None:
        """
        Log out of the current Power BI session and log back in.
        Used by Phase 2 / 3 to get a fresh session and clear any stale state.
        """
        log.info("  Re-logging in for fresh session...")
        # Clear cached URLs so workspace navigation runs again
        self._report_urls.clear()
        self._aom_trigger_idx = None
        page = self._page
        # Navigate to Microsoft logout
        try:
            page.goto("https://login.microsoftonline.com/logout.srf",
                      timeout=15_000, wait_until="domcontentloaded")
            page.wait_for_timeout(2_000)
        except Exception:
            pass
        # Re-run the login flow
        self._login()
        log.info("  Re-login complete.")

    def _discover_report_url(self, workspace: str, report_name: str,
                             known_url: str = None) -> str:
        """
        Navigate through the Power BI workspace to find the report and
        return the base report URL.

        Always uses workspace navigation so Power BI properly establishes
        a session and loads all report visuals before returning.
        """
        page = self._page

        log.info(f'  Looking for workspace "{workspace}"...')

        ws_href = None
        ws_selectors = [
            f'a[aria-label="{workspace}"]',
            f'a[title="{workspace}"]',
            f'a:has-text("{workspace}")',
        ]
        for sel in ws_selectors:
            try:
                el = page.locator(sel).first
                el.wait_for(timeout=5_000, state="attached")
                ws_href = el.get_attribute("href")
                if ws_href:
                    log.info(f'  Found workspace in sidebar.')
                    break
            except PWTimeout:
                continue

        if not ws_href:
            try:
                page.locator(
                    '[aria-label="Workspaces"], [data-testid="workspaces-nav-item"], '
                    'button:has-text("Workspaces")'
                ).first.click(timeout=8_000)
                page.wait_for_timeout(1_500)
            except PWTimeout:
                pass

            for sel in ws_selectors:
                try:
                    el = page.locator(sel).first
                    el.wait_for(timeout=8_000, state="attached")
                    ws_href = el.get_attribute("href")
                    if ws_href:
                        break
                except PWTimeout:
                    continue

        if not ws_href:
            os.makedirs(EXPORTS_DIR, exist_ok=True)
            page.screenshot(path=os.path.join(EXPORTS_DIR, "debug_workspace.png"))
            raise RuntimeError(
                f'Could not find workspace "{workspace}". '
                f'Screenshot saved to exports/debug_workspace.png'
            )

        ws_url = f"https://app.powerbi.com{ws_href}" if ws_href.startswith("/") else ws_href
        log.info(f'  Navigating to workspace directly...')
        page.goto(ws_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(4_000)   # wait for workspace content list to render
        self._handle_identity_prompt()  # dismiss verify-identity dialog if it appears
        self._dismiss_popups()

        # Session is now established via workspace navigation.
        # If the report URL is known from config, use it directly
        # (virtual-scroll list in workspace may not show all reports in DOM).
        if known_url:
            log.info(f'  Session established. Navigating to report via known URL...')
            page.goto(known_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(8_000)   # wait for all report visuals to initialise
            self._handle_identity_prompt()
            report_url = page.url.split("?")[0].rstrip("/")
            log.info(f'  Report URL: {report_url}')
            return report_url

        log.info(f'  Looking for report "{report_name}"...')

        report_href = None
        report_selectors = [
            f'a[aria-label="{report_name}"]',
            f'a[title="{report_name}"]',
            f'a:has-text("{report_name}")',
        ]
        for sel in report_selectors:
            try:
                el = page.locator(sel).first
                el.wait_for(timeout=10_000, state="attached")
                report_href = el.get_attribute("href")
                if report_href:
                    break
            except PWTimeout:
                continue

        if not report_href:
            os.makedirs(EXPORTS_DIR, exist_ok=True)
            page.screenshot(path=os.path.join(EXPORTS_DIR, "debug_report.png"))
            raise RuntimeError(
                f'Could not find report "{report_name}" in workspace "{workspace}". '
                f'Screenshot saved to exports/debug_report.png'
            )

        report_url_full = f"https://app.powerbi.com{report_href}" if report_href.startswith("/") else report_href
        log.info(f'  Navigating to report...')
        page.goto(report_url_full, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(8_000)   # wait for all report visuals to initialise

        self._handle_identity_prompt()

        report_url = page.url.split("?")[0].rstrip("/")
        log.info(f'  Report URL: {report_url}')
        return report_url

    # -- Export ----------------------------------------------------------------

    def export_report(self, filter_email: str, aom_name: str, date_str: str,
                      other_emails: list = None, report_cfg: dict = None) -> Optional[str]:
        """
        For this AOM:
          1. Discover report URL (cached; uses known URL from config if set)
          2. Navigate to base report URL (clean state, 4s wait)
          3. Open Filters pane
          4. Apply filter via slicer → Page 1 fallback → Filters pane card
          5. Smart-wait until conflicting AOM data disappears
          6. Verify only the expected AOM data is on screen
          7. Export as PDF

        Returns local PDF path on success, None on failure.
        """
        if not report_cfg:
            raise ValueError("report_cfg is required")

        workspace     = report_cfg["workspace"]
        report_name   = report_cfg["name"]
        filter_column = report_cfg.get("filter_column", "AOM")

        known_url = report_cfg.get("url")

        # Step 1: Discover report URL (cached after first call).
        # Always navigates through workspace first to establish a session,
        # then uses known_url if available to bypass virtual-scroll list.
        if (workspace, report_name) not in self._report_urls:
            self._report_urls[(workspace, report_name)] = self._discover_report_url(
                workspace, report_name, known_url=known_url
            )
        report_url = self._report_urls[(workspace, report_name)]

        page = self._page

        # Step 2: Navigate to base report URL (flush any residual filter)
        log.info(f"  Navigating to base report URL (clean state)...")
        page.goto(report_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(4_000)
        self._handle_identity_prompt()

        # Step 3: Open the Filters pane
        log.info(f"  Opening Filters pane...")
        self._open_filters_pane()
        page.wait_for_timeout(1_500)

        # Step 4: Set filter via UI
        log.info(f"  Setting filter: {filter_column} = {filter_email}")
        if not self._apply_filter_via_pane(filter_email, filter_column):
            # ── Inline retry: reload page and try once more ───────────────
            log.warning(f"  Filter failed on first attempt. Reloading and retrying...")
            page.wait_for_timeout(3_000)
            page.goto(report_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(5_000)
            self._handle_identity_prompt()
            self._open_filters_pane()
            page.wait_for_timeout(1_500)
            if not self._apply_filter_via_pane(filter_email, filter_column):
                log.error(
                    f"  FILTER APPLY FAILED \u2014 {aom_name}:\n"
                    f"  Could not set '{filter_column}' = '{filter_email}' in UI.\n"
                    f"  This report will NOT be attached."
                )
                self._debug_screenshot(page, f"{aom_name}_{report_name}")
                return None

        # Step 5: Smart wait – poll until other AOM emails are gone from the page
        MAX_WAIT_SEC  = 90
        POLL_INTERVAL = 5_000   # ms
        elapsed_sec   = 0

        log.info(f"  Polling until data refreshes (max {MAX_WAIT_SEC}s)...")
        while elapsed_sec < MAX_WAIT_SEC:
            page.wait_for_timeout(POLL_INTERVAL)
            elapsed_sec += POLL_INTERVAL // 1_000
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(200)
                page_text = page.evaluate("() => document.body.innerText")
                conflicts = [e for e in (other_emails or []) if e in page_text]
                if not conflicts:
                    log.info(
                        f"  Data looks clean after {elapsed_sec}s "
                        f"\u2014 no other AOM emails found. Proceeding."
                    )
                    break
                log.info(
                    f"  [{elapsed_sec}s] Still waiting \u2014 "
                    f"conflicting emails present: {', '.join(conflicts)}"
                )
            except Exception:
                pass

        if not self._verify_filter_on_screen(filter_email, f"{aom_name}_{report_name}", other_emails or []):
            return None

        # Step 6: Export
        safe_aom = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name).strip().replace(" ", "_")
        safe_rep = "".join(c if c.isalnum() or c in " _-" else "_" for c in report_name).strip().replace(" ", "_")
        fname = f"{safe_aom}_{safe_rep}_{date_str}.pdf"
        fpath = os.path.join(EXPORTS_DIR, fname)

        try:
            return self._trigger_pdf_export(page, fpath, f"{aom_name}_{report_name}")
        except Exception as e:
            log.error(f"  Export failed: {e}")
            self._debug_screenshot(page, f"{aom_name}_{report_name}")
            return None

    def _open_filters_pane(self) -> bool:
        """
        Ensure the Power BI Filters pane is open and visible.
        Checks if already open; if not, clicks the toggle button.
        Returns True if the pane is (or becomes) visible.
        """
        page = self._page

        # Check if already visible
        PANE_SELECTORS = [
            '[data-automation-id="filters-pane"]',
            '.filterExplorerContainer',
            '[aria-label="Filters pane"]',
            '.filtersPane',
            '.report-sidePane',
        ]
        for sel in PANE_SELECTORS:
            try:
                if page.locator(sel).first.is_visible(timeout=1_500):
                    log.info("  Filters pane already open.")
                    return True
            except Exception:
                continue

        # Not visible â€” click the toggle button
        TOGGLE_SELECTORS = [
            '[aria-label="Open Filters pane"]',
            '[aria-label="Filters"]',
            'button[title="Filters"]',
            '[data-testid="filters-pane-toggle"]',
            'button:has-text("Filters")',
        ]
        for sel in TOGGLE_SELECTORS:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=2_000):
                    btn.click(force=True)
                    page.wait_for_timeout(2_000)
                    log.info("  Filters pane opened via toggle button.")
                    return True
            except Exception:
                continue

        log.warning("  Could not confirm Filters pane is open â€” continuing anyway.")
        return False


    def _try_slicer(self, filter_email: str) -> bool:
        """
        Locate the AOM slicer by its title text 'AOM', clear any previous
        selection via the eraser button, then open the dropdown and select
        the target name.

        Click strategy: locator.click(force=True) generates CDP-level trusted
        events. Combined with keyboard shortcuts (Enter/Space), this is the
        most reliable way to open Power BI's Angular dropdown slicer.
        Active wait (wait_for_selector, 10 s) instead of fixed sleep.
        Up to 5 attempts. No generic slicer hunt.
        """
        page = self._page

        # Force-unhide visuals (including display:none)
        try:
            page.evaluate("""
            () => {
                document.querySelectorAll(
                    '.visual-container, .visualContainer, [class*=\"visual\"]'
                ).forEach(v => {
                    v.style.setProperty('visibility',     'visible', 'important');
                    v.style.setProperty('opacity',        '1',       'important');
                    v.style.setProperty('pointer-events', 'auto',    'important');
                    if (getComputedStyle(v).display === 'none')
                        v.style.setProperty('display', 'block', 'important');
                });
            }""")
            page.wait_for_timeout(500)
        except Exception:
            pass

        # JS: find the AOM slicer trigger (prefer .slicerDropdownMenu)
        FIND_TRIGGER_JS = r"""
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let node;
            while ((node = walker.nextNode())) {
                if (node.textContent.trim() !== 'AOM') continue;
                let el = node.parentElement;
                for (let i = 0; i < 10; i++) {
                    if (!el || el === document.body) break;
                    const menu = el.querySelector('.slicerDropdownMenu');
                    if (menu) return menu;
                    const combo = el.querySelector(
                        '[role="combobox"], [aria-haspopup="listbox"], [aria-haspopup="true"]'
                    );
                    if (combo) return combo;
                    el = el.parentElement;
                }
            }
            return null;
        }
        """

        # JS: find eraser/clear button near AOM text
        FIND_ERASER_JS = r"""
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let node;
            while ((node = walker.nextNode())) {
                if (node.textContent.trim() !== 'AOM') continue;
                let el = node.parentElement;
                for (let i = 0; i < 8; i++) {
                    if (!el || el === document.body) break;
                    const e = el.querySelector(
                        '[aria-label*="clear" i], [aria-label*="Clear" i], '
                        + '[aria-label*="eraser" i], [title*="clear" i], '
                        + '.slicerDeleteButton, [class*="clearButton"]'
                    );
                    if (e) return e;
                    el = el.parentElement;
                }
            }
            return null;
        }
        """

        OPTION_SELS = '[role="option"], [role="listbox"] li, div[role="listbox"] span'
        EMAIL_SELS  = [
            f'[role="option"]:has-text("{filter_email}")',
            f'[role="listbox"] li:has-text("{filter_email}")',
            f'div[role="listbox"] span:has-text("{filter_email}")',
        ]

        def _unhide_chain(elem_h):
            try:
                page.evaluate(r"""(node) => {
                    let c = node;
                    while (c && c !== document.body) {
                        c.style.setProperty('visibility',     'visible', 'important');
                        c.style.setProperty('opacity',        '1',       'important');
                        c.style.setProperty('pointer-events', 'auto',    'important');
                        if (getComputedStyle(c).display === 'none')
                            c.style.setProperty('display', 'block', 'important');
                        c = c.parentElement;
                    }
                }""", elem_h)
            except Exception:
                pass

        def _open_dropdown(elem_h):
            """Unhide chain, focus, keyboard, then force-click."""
            _unhide_chain(elem_h)
            page.wait_for_timeout(200)
            try:
                elem_h.focus()
                page.wait_for_timeout(200)
            except Exception:
                pass
            for key in ['Enter', 'Space', 'ArrowDown']:
                try:
                    page.keyboard.press(key)
                    page.wait_for_timeout(250)
                except Exception:
                    pass
            try:
                elem_h.click(force=True, timeout=3_000)
            except Exception:
                pass

        for attempt in range(5):
            log.info(f'  AOM slicer: attempt {attempt + 1}/5 for \'{filter_email}\'...')

            try:
                h    = page.evaluate_handle(FIND_TRIGGER_JS)
                elem = h.as_element()
            except Exception:
                elem = None

            if elem is None:
                log.warning('  AOM slicer trigger not found in DOM.')
                return False

            # Attempt 1: hover to reveal eraser, then click it
            if attempt == 0:
                try:
                    elem.hover(timeout=2_000)
                    page.wait_for_timeout(400)
                    eh     = page.evaluate_handle(FIND_ERASER_JS)
                    eraser = eh.as_element()
                    if eraser:
                        _unhide_chain(eraser)
                        eraser.click(force=True, timeout=2_000)
                        page.wait_for_timeout(800)
                        log.info('  Eraser clicked — previous selection cleared.')
                except Exception:
                    pass

            _open_dropdown(elem)

            # Active-wait for options (up to 10 s)
            try:
                page.wait_for_selector(OPTION_SELS, timeout=10_000, state='attached')
                options_found = True
            except Exception:
                options_found = False

            if not options_found:
                page.keyboard.press('Escape')
                page.wait_for_timeout(1_000)
                log.info(f'  No dropdown options in 10 s (attempt {attempt + 1}).')
                continue

            target_found = any(page.locator(s).count() > 0 for s in EMAIL_SELS)
            if not target_found:
                page.keyboard.press('Escape')
                page.wait_for_timeout(500)
                log.warning(
                    f'  Dropdown opened but \'{filter_email}\' not listed. ')
                return False

            log.info(f'  AOM dropdown: \'{filter_email}\' found \u2714')

            for sel in ['[role="option"]:has-text("Select all")',
                        '[role="option"]:has-text("(Select all)")']:
                try:
                    if page.locator(sel).count() > 0:
                        page.locator(sel).first.click(force=True, timeout=2_000)
                        page.wait_for_timeout(600)
                        log.info('  \'Select all\' deselected \u2713')
                        break
                except Exception:
                    pass

            for sel in EMAIL_SELS:
                try:
                    if page.locator(sel).count() > 0:
                        opt = page.locator(sel).first
                        opt.evaluate('el => el.click()')
                        page.wait_for_timeout(400)
                        opt.click(force=True, timeout=2_000)
                        page.wait_for_timeout(1_500)
                        log.info(f'  Selected \'{filter_email}\' in AOM slicer \u2713')
                        page.keyboard.press('Escape')
                        page.wait_for_timeout(600)
                        return True
                except Exception:
                    continue

            page.keyboard.press('Escape')
            return False

        log.warning(f'  AOM slicer: all 5 attempts failed for \'{filter_email}\'.')
        return False

    def _apply_filter_via_pane(self, filter_email: str, filter_column: str) -> bool:
        """
        Set the AOM filter through the Power BI UI (slicer on current page only).
        The AOM slicer is always on the first/current page — no page-switching needed.
        Falls back to the Filters pane card if the slicer attempt fails.
        """
        page = self._page

        # Primary: AOM slicer
        log.info("  Step A: Opening AOM slicer dropdown...")
        if self._try_slicer(filter_email):
            return True

        # Fallback: Filters pane card
        log.info("  Step B: Trying Filters pane card for 'AOM'...")
        log.info(f"  Step B: Trying Filters pane card for '{filter_column}'...")
        card_found = False
        try:
            card = page.locator(
                'filter-pane, .filterPane, [aria-label="Filters"]'
            ).get_by_text(filter_column, exact=True).first
            card.wait_for(state="visible", timeout=5_000)
            card.click(force=True)
            page.wait_for_timeout(1_500)
            card_found = True
            log.info(f"  Filter card '{filter_column}' expanded.")
        except Exception:
            pass

        if not card_found:
            log.error(f"  '{filter_column}' filter card not found.")
            return False

        for text in ["(Select all)", "Select all"]:
            try:
                el = page.get_by_text(text, exact=True).first
                if el.is_visible(timeout=2_000):
                    el.click(force=True)
                    page.wait_for_timeout(700)
                    log.info(f"  Deselected '{text}' in filter card \u2713")
                    break
            except Exception:
                continue

        for sel in [
            f'[role="option"]:has-text("{filter_email}")',
            f'label:has-text("{filter_email}")',
        ]:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=3_000):
                    el.click(force=True)
                    page.wait_for_timeout(1_500)
                    log.info(f"  Selected '{filter_email}' in filter card \u2713")
                    return True
            except Exception:
                continue

        log.error(
            f"  All approaches failed.\n"
            f"  Could not set '{filter_column}' = '{filter_email}'.\n"
            f"  Email will NOT be sent."
        )
        return False

    def _verify_filter_on_screen(self, filter_email: str, aom_name: str,
                                  other_emails: list) -> bool:
        """
        Screenshot the current report state and read all visible page text.

        Checks:
          1. Expected AOM email IS visible in the data â†’ proves filter is active
          2. No OTHER AOM's email is visible in the data â†’ proves no wrong data

        Returns True  â†’ safe to export
        Returns False â†’ conflict or cannot verify â†’ email is NOT sent, counts as FAIL
        Reason is always logged explicitly.
        """
        page = self._page
        page.wait_for_timeout(3_000)   # let all visuals fully render

        # Always take a screenshot (audit trail + debugging)
        # Strip + replace spaces: "Ritesh Soni " â†’ "verify_Ritesh_Soni.png"
        # Windows rejects filenames ending with space (Errno 22).
        safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name.strip())
        safe = safe.replace(" ", "_")
        os.makedirs(EXPORTS_DIR, exist_ok=True)
        shot_path = os.path.join(EXPORTS_DIR, f"verify_{safe}.png")

        try:
            page.screenshot(path=shot_path, full_page=False)
            log.info(f"  Verification screenshot: {os.path.basename(shot_path)}")
        except Exception as e:
            log.warning(f"  Screenshot failed: {e}")

        # Read all visible text from the page
        # First close any open slicer dropdown â€” unchecked options appear as text
        # and would be mistaken for conflicting data if not dismissed first.
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)
        except Exception:
            pass
        try:
            page_text = page.evaluate("""
                () => {
                    // Clone the body so we don't mutate the live DOM
                    let clone = document.body.cloneNode(true);
                    // Remove all slicer-related elements â€” they store all option
                    // names in the DOM even when the dropdown is closed, which
                    // would cause false "conflict" detections.
                    clone.querySelectorAll(
                        '.slicer-container, .visual-slicer, ' +
                        '[class*="slicer"], [class*="Slicer"], ' +
                        '.slicerDropdownMenu, [role="listbox"], ' +
                        '[aria-label*="slicer"], [aria-label*="Slicer"]'
                    ).forEach(el => el.remove());
                    return clone.innerText;
                }
            """)

        except Exception as e:
            log.error(
                f"  VERIFICATION FAILED â€” {aom_name}:\n"
                f"  Reason: Cannot read page content ({e})\n"
                f"  Email will NOT be sent â€” cannot confirm data is correct."
            )
            return False

        page_text_lower = page_text.lower()
        expected_lower  = filter_email.lower()
        others_lower    = [e.lower() for e in other_emails]

        # Check 1: any OTHER AOM's email visible in the report data?
        conflicting = [e for e in others_lower if e in page_text_lower]
        if conflicting:
            log.error(
                f"\n"
                f"  â•”â•â• DATA CONFLICT â€” EMAIL WILL NOT BE SENT â•â•â•â•â•â•â•—\n"
                f"  â•‘  AOM              : {aom_name}\n"
                f"  â•‘  Expected filter  : {filter_email}\n"
                f"  â•‘  Conflicting data : {', '.join(set(conflicting))}\n"
                f"  â•‘  Report contains another AOM's data. Aborting.\n"
                f"  â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•"
            )
            return False

        # Check 2: the expected email must be visible (proves filter worked)
        if expected_lower not in page_text_lower:
            log.error(
                f"  VERIFICATION FAILED â€” {aom_name}:\n"
                f"  Reason: '{filter_email}' not found in visible report data.\n"
                f"  The filter may not have been applied or data is not loaded.\n"
                f"  Email will NOT be sent."
            )
            return False

        log.info(
            f"  Data verification âœ“ â€” only {filter_email} found.\n"
            f"  No conflicting AOM data. Safe to export."
        )
        return True

    def _trigger_pdf_export(self, page, fpath: str, aom_name: str) -> str:
        """
        Drive the Power BI UI to export the report as PDF.

        Power BI sometimes downloads the PDF directly (Playwright download event)
        and sometimes opens it in a new browser tab.  This method handles both.

        Flow:
          dismiss popups → click Export (toolbar) → click PDF option →
          confirm dialog (if shown) → capture via download event OR new tab.
        """
        import time as _time

        # Track new pages that open during export
        new_pages_opened: list = []

        def _on_new_page(p):
            new_pages_opened.append(p)

        page.context.on('page', _on_new_page)

        try:
            # Step 1: Dismiss popups
            self._dismiss_popups()
            page.wait_for_timeout(500)

            # Step 2: Click Export toolbar button
            log.info("  Clicking Export toolbar button...")
            page.locator(
                '[aria-label="Export"], button:has-text("Export")'
            ).first.click(force=True, timeout=10_000)
            page.wait_for_timeout(1_000)

            log.info("  Selecting PDF and waiting for download...")
            os.makedirs(EXPORTS_DIR, exist_ok=True)

            try:
                with page.expect_download(timeout=EXPORT_TIMEOUT) as dl_info:
                    # Click PDF option
                    page.locator(
                        '[role="menuitem"]:has-text("PDF"), '
                        '[role="option"]:has-text("PDF"), '
                        'button:has-text("PDF"), a:has-text("PDF")'
                    ).first.click(force=True, timeout=8_000)
                    page.wait_for_timeout(1_000)

                    # Try confirmation dialog (multiple possible selectors)
                    CONFIRM_SELS = [
                        '[role="dialog"] button:has-text("Export")',
                        '.ms-Dialog button:has-text("Export")',
                        '[role="dialog"] button:has-text("Download")',
                        '.ms-Dialog button:has-text("Download")',
                        '[role="dialog"] button.ms-Button--primary',
                        '.ms-Dialog-actions button.ms-Button--primary',
                        'button[data-testid*="export-confirm"]',
                    ]
                    confirmed = False
                    for sel in CONFIRM_SELS:
                        try:
                            btn = page.locator(sel).first
                            btn.wait_for(timeout=3_000, state='visible')
                            btn.click(force=True)
                            log.info(f"  Export dialog confirmed.")
                            confirmed = True
                            break
                        except Exception:
                            continue
                    if not confirmed:
                        log.info("  No export confirmation dialog \u2014 download started directly.")

                dl = dl_info.value
                dl.save_as(fpath)
                log.info(f"  PDF saved: {os.path.basename(fpath)}")
                return fpath

            except PWTimeout:
                log.warning("  Download event timed out. Checking for PDF in new tab...")

                # Wait up to 30 s for a new page to appear
                deadline = _time.time() + 30
                while not new_pages_opened and _time.time() < deadline:
                    page.wait_for_timeout(1_000)

                for np in new_pages_opened:
                    try:
                        np.wait_for_load_state('networkidle', timeout=30_000)
                        pdf_url = np.url
                        log.info(f"  New tab URL: {pdf_url[:100]}")

                        # Try downloading via HTTP using the browser's cookies
                        try:
                            import requests as _req
                            cookies = {
                                c['name']: c['value']
                                for c in page.context.cookies()
                            }
                            resp = _req.get(pdf_url, cookies=cookies,
                                            timeout=120, stream=True)
                            ct = resp.headers.get('content-type', '')
                            if resp.status_code == 200:
                                with open(fpath, 'wb') as f:
                                    for chunk in resp.iter_content(chunk_size=8192):
                                        f.write(chunk)
                                np.close()
                                if os.path.getsize(fpath) > 1_000:
                                    log.info(f"  PDF saved from new tab: "
                                             f"{os.path.basename(fpath)}")
                                    return fpath
                        except Exception as req_err:
                            log.warning(f"  HTTP download from new tab failed: {req_err}")

                        np.close()
                    except Exception as tab_err:
                        log.warning(f"  New tab handling failed: {tab_err}")
                        try:
                            np.close()
                        except Exception:
                            pass

                raise TimeoutError(
                    f"PDF export failed for '{aom_name}': "
                    "no download event and no PDF in new tab."
                )
        finally:
            try:
                page.context.remove_listener('page', _on_new_page)
            except Exception:
                pass


    def _debug_screenshot(self, page, aom_name: str) -> None:
        """Save a screenshot to exports/ to help debug failures."""
        try:
            safe  = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name)
            spath = os.path.join(EXPORTS_DIR, f"debug_{safe}.png")
            page.screenshot(path=spath)
            log.info(f"  Debug screenshot: {os.path.basename(spath)}")
        except Exception:
            pass

