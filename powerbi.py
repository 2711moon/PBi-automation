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
        self._slicer_cache = {}   # cache: report_name -> slicer index that contains AOM names

    # -- Context manager -------------------------------------------------------

    def __enter__(self) -> "PowerBIExporter":
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=True,
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
        Apply the AOM filter by hunting through slicer triggers on the page.

        For each slicer trigger found:
          1. Force-unhide parent chain via JS
          2. Open dropdown (keyboard shortcuts + force click)
          3. Check page-wide for [role="option"] elements
          4. If the target name is present, select it and return True
        """
        page = self._page

        # 1. Force-unhide all visuals via JS
        try:
            page.evaluate("""() => {
                const visuals = document.querySelectorAll(
                    '.visual-container, .visualContainer, [class*="visual"]'
                );
                visuals.forEach(v => {
                    v.style.setProperty('visibility', 'visible', 'important');
                    v.style.setProperty('opacity', '1', 'important');
                    v.style.setProperty('pointer-events', 'auto', 'important');
                    if (v.style.display === 'none') {
                        v.style.setProperty('display', 'block', 'important');
                    }
                });
            }""")
            page.wait_for_timeout(1_000)
        except Exception as e:
            log.warning(f"  JS unhide failed (ignoring): {e}")

        # 2. Find all potential slicer triggers
        SLICER_TRIGGERS = [
            '.visual [role="combobox"]',
            '.visual [aria-haspopup="listbox"]',
            '.visual [aria-haspopup="true"]',
            '.slicerDropdownMenu',
        ]

        potential_triggers = []
        for sel in SLICER_TRIGGERS:
            try:
                page.locator(sel).first.wait_for(state="attached", timeout=2_000)
                elements = page.locator(sel).all()
                potential_triggers.extend(elements)
            except Exception:
                pass

        if not potential_triggers:
            log.warning("  No slicer triggers found in DOM.")
            return False

        log.info(f"  Found {len(potential_triggers)} potential slicer(s). Hunting for '{filter_email}'...")

        target_slicer = None

        # 3. Hunt for the correct slicer
        for idx, el in enumerate(potential_triggers):
            try:
                # Bring element to front (unhide parent chain)
                page.evaluate("""(node) => {
                    let curr = node;
                    while (curr && curr !== document.body) {
                        curr.style.setProperty('z-index', '2147483647', 'important');
                        curr.style.setProperty('opacity', '1', 'important');
                        curr.style.setProperty('visibility', 'visible', 'important');
                        curr.style.setProperty('pointer-events', 'auto', 'important');
                        curr = curr.parentElement;
                    }
                }""", el)
                page.wait_for_timeout(300)

                # Open dropdown — keyboard shortcuts first, then mouse click
                try:
                    el.focus()
                    page.wait_for_timeout(200)
                    page.keyboard.press("Alt+ArrowDown")
                    page.wait_for_timeout(300)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(300)
                    page.keyboard.press("Space")
                except Exception as e:
                    log.warning(f"  Keyboard focus failed: {e}")

                el.click(force=True, timeout=2_000)
                page.wait_for_timeout(2_000)

                # Check if any options rendered page-wide
                option_sels = '[role="option"], [role="listbox"] li, div[role="listbox"] span'
                if page.locator(option_sels).count() == 0:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                    continue

                # Look for our specific name
                email_sels = [
                    f'[role="option"]:has-text("{filter_email}")',
                    f'[role="listbox"] li:has-text("{filter_email}")',
                    f'div[role="listbox"] span:has-text("{filter_email}")'
                ]

                found_email = False
                for sel in email_sels:
                    if page.locator(sel).count() > 0:
                        found_email = True
                        break

                if found_email:
                    log.info(f"  Found the correct slicer containing '{filter_email}'.")
                    target_slicer = el
                    break
                else:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
            except Exception:
                continue

        if not target_slicer:
            log.warning("  None of the slicers contained the target name.")
            return False

        # 4. Right slicer is open. Apply the filter.
        # Deselect "Select all" first
        for sel in [
            '[role="option"]:has-text("Select all")',
            '[role="option"]:has-text("(Select all)")'
        ]:
            try:
                if page.locator(sel).count() > 0:
                    page.locator(sel).first.click(force=True, timeout=2_000)
                    page.wait_for_timeout(800)
                    log.info("  'Select all' deselected \u2713")
                    break
            except Exception:
                continue

        # Select the specific AOM name
        email_sels = [
            f'[role="option"]:has-text("{filter_email}")',
            f'[role="listbox"] li:has-text("{filter_email}")',
            f'div[role="listbox"] span:has-text("{filter_email}")'
        ]

        email_clicked = False
        for sel in email_sels:
            try:
                if page.locator(sel).count() > 0:
                    email_opt = page.locator(sel).first
                    try:
                        email_opt.evaluate("el => el.click()")
                        page.wait_for_timeout(500)
                    except Exception:
                        pass
                    try:
                        email_opt.click(force=True, timeout=2_000)
                    except Exception:
                        pass
                    page.wait_for_timeout(1_500)
                    log.info(f"  Selected '{filter_email}' in slicer \u2713")
                    email_clicked = True
                    break
            except Exception:
                continue

        if not email_clicked:
            log.warning(f"  Failed to click email option.")
            page.keyboard.press("Escape")
            return False

        # Close the dropdown
        page.keyboard.press("Escape")
        page.wait_for_timeout(800)
        try:
            option_sels = '[role="option"], [role="listbox"] li, div[role="listbox"] span'
            if page.locator(option_sels).first.is_visible(timeout=600):
                log.info("  Dropdown still open \u2014 clicking trigger to close...")
                target_slicer.click(force=True, timeout=1_000)
                page.wait_for_timeout(600)
        except Exception:
            pass
        return True

    def _apply_filter_via_pane(self, filter_email: str, filter_column: str) -> bool:
        """
        Set the AOM filter through the Power BI UI.
        Flow: slicer on current page → Page 1 slicer fallback → Filters pane card.
        """
        page = self._page
        original_url = page.url

        # Primary: slicer on current page
        log.info("  Step A: Opening slicer dropdown on current page...")
        if self._try_slicer(filter_email):
            return True

        # Fallback 1: Hidden slicer on Page 1
        log.info("  Slicer not found on current page. Trying Page 1 fallback...")
        try:
            page.locator(
                'button[aria-label^="Page 1"], .pageNavigation button, .explorationContainer .navigation-node'
            ).first.click(timeout=5_000)
            log.info("  Navigated to Page 1.")
            page.wait_for_timeout(5_000)

            if self._try_slicer(filter_email):
                log.info("  Successfully filtered via Page 1 slicer. Navigating back...")
                page.goto(original_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                page.wait_for_timeout(4_000)
                return True
            else:
                log.info("  Slicer not found or failed on Page 1. Navigating back...")
                page.goto(original_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                page.wait_for_timeout(4_000)
        except Exception as e:
            log.warning(f"  Page 1 fallback failed: {e}")

        # Fallback 2: Filters pane filter card
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

        The Export button is a standalone button in the toolbar â€”
        NOT inside the File menu (File menu has no Export option here).

        Flow: dismiss popups â†’ click Export (toolbar) â†’ click PDF â†’
              confirm dialog if shown â†’ wait for download.

        IMPORTANT: page.expect_download() must be open BEFORE the action
        that triggers the download, not after. Both the PDF click and the
        confirm-dialog click are wrapped inside expect_download so whichever
        one starts the file download is captured correctly.
        """

        # Step 1: Dismiss Copilot dialog, CDK overlays, any popups
        self._dismiss_popups()
        page.wait_for_timeout(500)

        # Step 2: Click the Export toolbar button
        log.info("  Clicking Export toolbar button...")
        page.locator(
            '[aria-label="Export"], '
            'button:has-text("Export")'
        ).first.click(force=True, timeout=10_000)
        page.wait_for_timeout(800)

        # Step 3 + 4: Open expect_download FIRST, then click PDF and handle dialog.
        # Power BI either downloads immediately on PDF click, or shows a dialog first.
        # expect_download captures whichever action triggers the actual download.
        log.info("  Selecting PDF and waiting for download...")
        os.makedirs(EXPORTS_DIR, exist_ok=True)

        with page.expect_download(timeout=EXPORT_TIMEOUT) as dl_info:
            # Click PDF in the dropdown
            page.locator(
                '[role="menuitem"]:has-text("PDF"), '
                '[role="option"]:has-text("PDF"), '
                'button:has-text("PDF"), '
                'a:has-text("PDF")'
            ).first.click(force=True, timeout=8_000)
            page.wait_for_timeout(1_000)

            # If Power BI shows an export options/confirmation dialog, click Export in it
            try:
                confirm_btn = page.locator(
                    '[role="dialog"] button:has-text("Export"), '
                    '.ms-Dialog button:has-text("Export"), '
                    'button[data-testid*="export-confirm"]'
                ).first
                confirm_btn.wait_for(timeout=6_000, state="visible")
                confirm_btn.click(force=True)
                log.info("  Export dialog confirmed.")
            except PWTimeout:
                log.info("  No export confirmation dialog â€” download started directly.")

        dl = dl_info.value
        dl.save_as(fpath)
        log.info(f"  PDF saved: {os.path.basename(fpath)}")
        return fpath


    def _debug_screenshot(self, page, aom_name: str) -> None:
        """Save a screenshot to exports/ to help debug failures."""
        try:
            safe  = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name)
            spath = os.path.join(EXPORTS_DIR, f"debug_{safe}.png")
            page.screenshot(path=spath)
            log.info(f"  Debug screenshot: {os.path.basename(spath)}")
        except Exception:
            pass

