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

from config import (
    PBI_WORKSPACE_NAME, PBI_REPORT_NAME,
    PBI_FILTER_TABLE, PBI_FILTER_COLUMN,
    EXPORTS_DIR
)

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
        self.username    = username
        self.password    = password
        self._pw         = None
        self._browser    = None
        self._context    = None
        self._page       = None
        self._report_url = None   # discovered after first login

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

        log.info(f'Navigating to workspace "{PBI_WORKSPACE_NAME}" -> report "{PBI_REPORT_NAME}"...')
        self._report_url = self._discover_report_url()
        log.info(f"Report URL acquired: {self._report_url}")

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

        # Go directly to Microsoft login — skips the app.powerbi.com SSO redirect chain
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
        # (Microsoft may redirect to M365 home — this ensures we land on Power BI)
        log.info("  Navigating to Power BI...")
        page.goto("https://app.powerbi.com/", timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(3_000)

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
        Power BI sometimes shows 'You need to verify your identity' dialog
        when opening a report. It can appear twice back-to-back.
        Click 'Continue' each time it appears (up to 2 times).
        """
        for attempt in range(1, 3):
            try:
                btn = self._page.locator(
                    'button:has-text("Continue"), '
                    '[aria-label="Continue"]'
                ).first
                btn.wait_for(timeout=5_000, state="visible")
                btn.click()
                log.info(f"  Identity prompt dismissed (attempt {attempt}).")
                self._page.wait_for_timeout(2_000)
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

    def _discover_report_url(self) -> str:

        """
        Find the workspace and report by name, then return the base report URL.

        Instead of clicking (which can be blocked by overlay modals),
        we extract the href from the link and navigate directly via page.goto().
        """
        page = self._page

        # Dismiss any welcome/tour/modal dialogs that Power BI shows after login
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(800)
        except Exception:
            pass

        # -- Step 1: Find the workspace link and extract its URL ---------------
        log.info(f'  Looking for workspace "{PBI_WORKSPACE_NAME}"...')

        # Try the sidebar first (workspace is often pinned there)
        ws_href = None
        ws_selectors = [
            f'a[aria-label="{PBI_WORKSPACE_NAME}"]',
            f'a[title="{PBI_WORKSPACE_NAME}"]',
            f'a:has-text("{PBI_WORKSPACE_NAME}")',
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
            # Workspace not visible in sidebar — open the Workspaces panel
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
                f'Could not find workspace "{PBI_WORKSPACE_NAME}". '
                f'Screenshot saved to exports/debug_workspace.png'
            )

        # Navigate directly to the workspace URL (bypasses any overlay blocking clicks)
        # Use domcontentloaded — Power BI is a SPA and never reaches networkidle
        ws_url = f"https://app.powerbi.com{ws_href}" if ws_href.startswith("/") else ws_href
        log.info(f'  Navigating to workspace directly...')
        page.goto(ws_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(4_000)   # wait for workspace content list to render

        # -- Step 2: Find the report link and extract its URL ------------------
        log.info(f'  Looking for report "{PBI_REPORT_NAME}"...')

        report_href = None
        report_selectors = [
            f'a[aria-label="{PBI_REPORT_NAME}"]',
            f'a[title="{PBI_REPORT_NAME}"]',
            f'a:has-text("{PBI_REPORT_NAME}")',
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
                f'Could not find report "{PBI_REPORT_NAME}" in workspace "{PBI_WORKSPACE_NAME}". '
                f'Screenshot saved to exports/debug_report.png'
            )

        # Navigate directly to the report
        report_url_full = f"https://app.powerbi.com{report_href}" if report_href.startswith("/") else report_href
        log.info(f'  Navigating to report...')
        page.goto(report_url_full, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(8_000)   # wait for all report visuals to initialise

        # Dismiss "You need to verify your identity" dialog if it appears (up to 2x)
        self._handle_identity_prompt()

        # Extract clean base URL (strip any query params)
        report_url = page.url.split("?")[0].rstrip("/")
        log.info(f'  Report URL: {report_url}')
        return report_url

    # -- Export ----------------------------------------------------------------

    def _build_filter_url(self, filter_value: str) -> str:
        """
        Append an OData URL filter to the report URL for the given AOM email.

        Power BI URL filter spec:
          - Spaces in table/column NAMES must be encoded as '_x0020_' (OData encoding)
            NOT '%20' — Power BI silently ignores filters with %20 in identifiers.
          - The email value goes inside single quotes as-is (no encoding needed).

        Example output:
          ?filter=Store_x0020_Master/AOM_x0020_Mail_x0020_Id eq 'aomnorth1@kisna.com'
        """
        table  = PBI_FILTER_TABLE.replace(" ", "_x0020_")
        column = PBI_FILTER_COLUMN.replace(" ", "_x0020_")
        expr   = f"{table}/{column} eq '{filter_value}'"
        filter_url = f"{self._report_url}?filter={expr}"
        log.info(f"  Filter URL: ...?filter={table}/{column} eq '{filter_value}'")
        return filter_url


    def export_report(self, filter_email: str, aom_name: str, date_str: str,
                      other_emails: list = None) -> Optional[str]:
        """
        For this AOM:
          1. Navigate to the base report URL (clean state)
          2. Set the AOM Mail Id filter directly via the Filters pane UI
          3. Screenshot + verify only the expected AOM's data is on screen
          4. Export as PDF

        Returns local PDF path on success.
        Returns None on failure — main.py counts as FAIL, email is NOT sent.
        """
        page = self._page

        # ── Step 1: Navigate to base report URL (flush any residual filter) ───
        log.info(f"  Navigating to base report URL (clean state)...")
        page.goto(self._report_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(4_000)
        self._handle_identity_prompt()

        # ── Step 2: Open the Filters pane ─────────────────────────────────────
        log.info(f"  Opening Filters pane...")
        self._open_filters_pane()
        page.wait_for_timeout(1_500)

        # ── Step 3: Set filter via Filters pane UI ────────────────────────────
        log.info(f"  Setting filter: {PBI_FILTER_COLUMN} = {filter_email}")
        if not self._apply_filter_via_pane(filter_email):
            log.error(
                f"  FILTER APPLY FAILED — {aom_name}:\n"
                f"  Could not set '{PBI_FILTER_COLUMN}' = '{filter_email}' in Filters pane.\n"
                f"  Email will NOT be sent."
            )
            self._debug_screenshot(page, aom_name)
            return None

        # ── Smart wait: poll until other AOM emails are gone from the page ────
        # Power BI fires an async DAX query after a slicer change — the data table
        # can take anywhere from 3 to 30+ seconds to refresh depending on server load.
        # We poll every 5 seconds and exit as soon as the data looks clean.
        # Max wait: 90 seconds — after that we proceed and let verify() decide.
        MAX_WAIT_SEC  = 90
        POLL_INTERVAL = 5_000   # ms
        elapsed_sec   = 0

        log.info(f"  Polling until data refreshes (max {MAX_WAIT_SEC}s)...")
        while elapsed_sec < MAX_WAIT_SEC:
            page.wait_for_timeout(POLL_INTERVAL)
            elapsed_sec += POLL_INTERVAL // 1_000
            try:
                # ── Close slicer dropdown before reading page text ─────────────
                # The dropdown shows ALL option values (including unchecked ones).
                # If it's open, innerText will contain other AOM emails that are
                # just unchecked options — NOT actual data — causing false conflicts.
                page.keyboard.press("Escape")
                page.wait_for_timeout(200)

                page_text = page.evaluate("() => document.body.innerText")
                conflicts = [e for e in (other_emails or []) if e in page_text]
                if not conflicts:
                    log.info(
                        f"  Data looks clean after {elapsed_sec}s "
                        f"— no other AOM emails found. Proceeding."
                    )
                    break
                log.info(
                    f"  [{elapsed_sec}s] Still waiting — "
                    f"conflicting emails present: {', '.join(conflicts)}"
                )
            except Exception:
                pass   # page eval failed — keep waiting


        if not self._verify_filter_on_screen(filter_email, aom_name, other_emails or []):
            return None   # fail logged inside; main.py: fail += 1, email skipped

        # ── Step 6: Export ────────────────────────────────────────────────────
        safe  = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name)
        fname = f"{safe.replace(' ', '_')}_{date_str}.pdf"
        fpath = os.path.join(EXPORTS_DIR, fname)

        try:
            return self._trigger_pdf_export(page, fpath, aom_name)
        except Exception as e:
            log.error(f"  Export failed: {e}")
            self._debug_screenshot(page, aom_name)
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

        # Not visible — click the toggle button
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

        log.warning("  Could not confirm Filters pane is open — continuing anyway.")
        return False


    def _apply_filter_via_pane(self, filter_email: str) -> bool:
        """
        Set the AOM Mail Id filter directly through the Power BI Filters pane UI.

        Interacts with the AOM Mail Id slicer dropdown visual on the canvas.

        Flow:
          1. Click the slicer dropdown (showing "All") to open it
          2. Deselect "Select all" — removes the "All" selection completely
          3. Click ONLY the specific AOM email from the items list
          4. Close the dropdown

        Falls back to the Filters pane filter card if the slicer approach fails.
        Uses [role="option"] selectors throughout — this is safe because:
          - Slicer/filter items have role="option"
          - Data table cells have role="gridcell" → they will NEVER be clicked
        """
        page = self._page

        # ── Primary: interact with the slicer dropdown on the canvas ──────────
        log.info("  Step A: Opening AOM Mail Id slicer dropdown...")

        SLICER_TRIGGERS = [
            '[role="combobox"]',
            '[aria-haspopup="listbox"]',
            '[aria-haspopup="true"]',
            '.slicerDropdownMenu',
            '.slicerDropdown [tabindex]',
        ]
        slicer_opened = False
        for sel in SLICER_TRIGGERS:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=2_000):
                    el.click(force=True)
                    page.wait_for_timeout(1_200)
                    slicer_opened = True
                    log.info(f"  Slicer dropdown opened (selector: {sel})")
                    break
            except Exception:
                continue

        if slicer_opened:
            # ── Wait for dropdown items to load ───────────────────────────────
            # IMPORTANT: when there's no identity prompt, the report loads faster
            # and the slicer dropdown items may not be in the DOM yet when we
            # immediately search. We must wait for them to appear first.
            try:
                page.wait_for_selector('[role="option"]', timeout=8_000)
                log.info("  Slicer items loaded and ready.")
            except Exception:
                log.warning("  [role='option'] items not found after 8s — trying anyway...")

            # ── Deselect "Select all" FIRST ──────────────────────────────────
            # Power BI slicer: when "All" is selected, we must uncheck "Select all"
            # before checking a single value — otherwise the single value is ignored.
            DESELECT_SELS = [
                '[role="option"]:has-text("Select all")',
                '[role="option"]:has-text("(Select all)")',
                '[role="listbox"] li:has-text("Select all")',
                'div[role="listbox"] span:has-text("Select all")',
            ]
            deselected = False
            for sel in DESELECT_SELS:
                try:
                    el = page.locator(sel).first
                    if el.is_visible(timeout=2_000):
                        el.click(force=True)
                        page.wait_for_timeout(700)
                        deselected = True
                        log.info(f"  'Select all' deselected ✓")
                        break
                except Exception:
                    continue
            if not deselected:
                log.warning("  'Select all' not found in slicer — may not be needed.")

            # ── Select the specific AOM email ─────────────────────────────────
            EMAIL_SELS = [
                f'[role="option"]:has-text("{filter_email}")',
                f'[role="listbox"] li:has-text("{filter_email}")',
                f'div[role="listbox"] span:has-text("{filter_email}")',
            ]
            for sel in EMAIL_SELS:
                try:
                    el = page.locator(sel).first
                    if el.is_visible(timeout=8_000):   # ← increased from 3s to 8s
                        el.click(force=True)
                        page.wait_for_timeout(1_500)
                        log.info(f"  Selected '{filter_email}' in slicer ✓")
                        # Close the dropdown — Escape first, then confirm it's closed.
                        # Power BI's custom dropdown may not respond to Escape alone,
                        # so we check and click the trigger again if still open.
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(800)
                        try:
                            if page.locator('[role="option"]').first.is_visible(timeout=600):
                                log.info("  Dropdown still open — clicking trigger to close...")
                                page.locator('[role="combobox"]').first.click(force=True)
                                page.wait_for_timeout(600)
                        except Exception:
                            pass
                        return True
                except Exception:
                    continue

            log.warning("  Could not select email from slicer dropdown. Trying Filters pane...")

        # ── Fallback: Filters pane filter card ────────────────────────────────
        log.info(f"  Step B: Trying Filters pane card for '{PBI_FILTER_COLUMN}'...")

        card_found = False
        try:
            card = page.get_by_text(PBI_FILTER_COLUMN, exact=True).first
            card.wait_for(state="visible", timeout=5_000)
            card.click(force=True)
            page.wait_for_timeout(1_500)
            card_found = True
            log.info(f"  Filter card '{PBI_FILTER_COLUMN}' expanded.")
        except Exception:
            pass

        if not card_found:
            log.error(f"  '{PBI_FILTER_COLUMN}' filter card not found.")
            return False

        # Deselect "(Select all)" in the filter card
        for text in ["(Select all)", "Select all"]:
            try:
                el = page.get_by_text(text, exact=True).first
                if el.is_visible(timeout=2_000):
                    el.click(force=True)
                    page.wait_for_timeout(700)
                    log.info(f"  Deselected '{text}' in filter card ✓")
                    break
            except Exception:
                continue

        # Select the email using role="option" — safe, won't hit data table
        for sel in [
            f'[role="option"]:has-text("{filter_email}")',
            f'label:has-text("{filter_email}")',
        ]:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=3_000):
                    el.click(force=True)
                    page.wait_for_timeout(1_500)
                    log.info(f"  Selected '{filter_email}' in filter card ✓")
                    return True
            except Exception:
                continue

        log.error(
            f"  All approaches failed.\n"
            f"  Could not set '{PBI_FILTER_COLUMN}' = '{filter_email}'.\n"
            f"  Email will NOT be sent."
        )
        return False


    def _verify_filter_on_screen(self, filter_email: str, aom_name: str,
                                  other_emails: list) -> bool:
        """
        Screenshot the current report state and read all visible page text.

        Checks:
          1. Expected AOM email IS visible in the data → proves filter is active
          2. No OTHER AOM's email is visible in the data → proves no wrong data

        Returns True  → safe to export
        Returns False → conflict or cannot verify → email is NOT sent, counts as FAIL
        Reason is always logged explicitly.
        """
        page = self._page
        page.wait_for_timeout(3_000)   # let all visuals fully render

        # Always take a screenshot (audit trail + debugging)
        # Strip + replace spaces: "Ritesh Soni " → "verify_Ritesh_Soni.png"
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
        # First close any open slicer dropdown — unchecked options appear as text
        # and would be mistaken for conflicting data if not dismissed first.
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)
        except Exception:
            pass
        try:
            page_text = page.evaluate("() => document.body.innerText")

        except Exception as e:
            log.error(
                f"  VERIFICATION FAILED — {aom_name}:\n"
                f"  Reason: Cannot read page content ({e})\n"
                f"  Email will NOT be sent — cannot confirm data is correct."
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
                f"  ╔══ DATA CONFLICT — EMAIL WILL NOT BE SENT ══════╗\n"
                f"  ║  AOM              : {aom_name}\n"
                f"  ║  Expected filter  : {filter_email}\n"
                f"  ║  Conflicting data : {', '.join(set(conflicting))}\n"
                f"  ║  Report contains another AOM's data. Aborting.\n"
                f"  ╚═════════════════════════════════════════════════╝"
            )
            return False

        # Check 2: the expected email must be visible (proves filter worked)
        if expected_lower not in page_text_lower:
            log.error(
                f"  VERIFICATION FAILED — {aom_name}:\n"
                f"  Reason: '{filter_email}' not found in visible report data.\n"
                f"  The filter may not have been applied or data is not loaded.\n"
                f"  Email will NOT be sent."
            )
            return False

        log.info(
            f"  Data verification ✓ — only {filter_email} found.\n"
            f"  No conflicting AOM data. Safe to export."
        )
        return True

    def _trigger_pdf_export(self, page, fpath: str, aom_name: str) -> str:
        """
        Drive the Power BI UI to export the report as PDF.

        The Export button is a standalone button in the toolbar —
        NOT inside the File menu (File menu has no Export option here).

        Flow: dismiss popups → click Export (toolbar) → click PDF →
              confirm dialog if shown → wait for download.

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
            page.wait_for_timeout(1_200)

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
                log.info("  No export confirmation dialog — download started directly.")

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
