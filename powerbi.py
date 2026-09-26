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

    def _discover_report_url(self, workspace: str, report_name: str) -> str:
        """
        Find the workspace and report by name, then return the base report URL.
        """
        page = self._page
        
        log.info("  Ensuring clean Power BI home state...")
        page.goto("https://app.powerbi.com/home", timeout=NAV_TIMEOUT)
        
        try:
            page.locator("nav, [data-testid='left-nav-pane']").first.wait_for(timeout=15_000, state="visible")
        except Exception:
            pass

        self._dismiss_popups()

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
            # Workspace not visible in sidebar â€” open the Workspaces panel
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

        # Navigate directly to the workspace URL (bypasses any overlay blocking clicks)
        # Use domcontentloaded â€” Power BI is a SPA and never reaches networkidle
        ws_url = f"https://app.powerbi.com{ws_href}" if ws_href.startswith("/") else ws_href
        log.info(f'  Navigating to workspace directly...')
        page.goto(ws_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(4_000)   # wait for workspace content list to render

        self._dismiss_popups()
        self._handle_identity_prompt()

        # -- Step 2: Find the report link and extract its URL ------------------
        log.info(f'  Looking for report "{report_name}"...')

        report_href = None
        report_selectors = [
            f'a[aria-label="{report_name}"]',
            f'a[title="{report_name}"]',
            f'a:has-text("{report_name}")',
        ]
        
        # 2a. Try to find the report link immediately
        for sel in report_selectors:
            try:
                el = page.locator(sel).first
                el.wait_for(timeout=3_000, state="attached")
                report_href = el.get_attribute("href")
                if report_href:
                    break
            except PWTimeout:
                continue
                
        # 2b. If not found, use the workspace "Filter by keyword" bar or the global search.
        if not report_href:
            try:
                # Dismiss any identity prompt that may be blocking the workspace
                self._handle_identity_prompt()

                # Try the "Filter by keyword" bar first (right side of workspace)
                filter_bar = page.locator('input[placeholder="Filter by keyword"], input[aria-label="Filter by keyword"]').first
                if filter_bar.is_visible(timeout=2_000):
                    filter_bar.fill(report_name)
                    page.wait_for_timeout(2_000)
                    log.info(f'  Used keyword filter to search for "{report_name}".')
                    for sel in report_selectors:
                        try:
                            el = page.locator(sel).first
                            el.wait_for(timeout=4_000, state="attached")
                            report_href = el.get_attribute("href")
                            if report_href:
                                break
                        except PWTimeout:
                            continue
            except Exception:
                pass

        if not report_href:
            try:
                # Fall back to global search bar autocomplete
                search_box = page.locator('input[placeholder="Search"], input[aria-label="Search"]').first
                if search_box.is_visible(timeout=2_000):
                    search_box.fill(report_name)
                    page.wait_for_timeout(2_000)
                    log.info(f'  Used global search for "{report_name}".')
                    
                    autocomplete_sel = f'[role="option"]:has-text("{report_name}"), [role="listitem"]:has-text("{report_name}")'
                    try:
                        result = page.locator(autocomplete_sel).first
                        result.wait_for(timeout=4_000, state="visible")
                        href_val = result.get_attribute("href")
                        if href_val and "/reports/" in href_val:
                            report_href = href_val
                            log.info("  Found report link in autocomplete href.")
                        else:
                            # Click and check where we land
                            result.click()
                            page.wait_for_timeout(5_000)
                            self._handle_identity_prompt()
                            landed = page.url
                            if "/reports/" in landed:
                                report_href = landed
                                log.info("  Found report link via autocomplete click.")
                            else:
                                # Still on workspace list â€” find the report link in the filtered list
                                for sel in report_selectors:
                                    try:
                                        el = page.locator(sel).first
                                        el.wait_for(timeout=4_000, state="attached")
                                        href_val = el.get_attribute("href")
                                        if href_val and "/reports/" in href_val:
                                            report_href = href_val
                                            log.info("  Found report link in filtered workspace list.")
                                            break
                                    except PWTimeout:
                                        continue
                    except PWTimeout:
                        pass
            except Exception:
                pass
                
        if report_href:
            log.info("  Found report link.")

        if not report_href:
            os.makedirs(EXPORTS_DIR, exist_ok=True)
            page.screenshot(path=os.path.join(EXPORTS_DIR, "debug_report.png"))
            raise RuntimeError(
                f'Could not find report "{report_name}" in workspace "{workspace}". '
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


    def export_report(self, filter_email: str, aom_name: str, date_str: str,
                      other_emails: list = None, report_cfg: dict = None) -> Optional[str]:
        """
        For this AOM:
          1. Discover/fetch the report URL
          2. Navigate to the base report URL
          3. Dismiss identity dialogs
          4. Navigate to Page 1 (slicer is a hidden visual on Page 1, synced across all pages)
          5. Force-unhide the slicer visual via JS and apply the AOM filter
          6. Verify data, then export as PDF

        Returns local PDF path on success, None on failure.
        """
        if not report_cfg:
            raise ValueError("report_cfg is required")

        workspace   = report_cfg["workspace"]
        report_name = report_cfg["name"]
        filter_column = report_cfg["filter_column"]

        # â”€â”€ Step 1: Discover report URL (cached after first run) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Step 1: Get report URL (from config or discover and cache)
        if "url" in report_cfg and report_cfg["url"]:
            report_url = report_cfg["url"]
        else:
            if (workspace, report_name) not in self._report_urls:
                self._report_urls[(workspace, report_name)] = self._discover_report_url(workspace, report_name)
            report_url = self._report_urls[(workspace, report_name)]
        page = self._page

        # â”€â”€ Step 2: Navigate to base report URL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log.info(f"  Navigating to base report URL (clean state)...")
        page.goto(report_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(5_000)

        # â”€â”€ Step 3: Dismiss identity dialog â€” aggressively â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        self._handle_identity_prompt()
        self._dismiss_popups()
        self._handle_identity_prompt()   # can appear twice

        # â”€â”€ Step 4: Navigate to Page 1 (slicer lives here, syncs all pages) â”€â”€
        log.info("  Navigating to Page 1 (slicer is on Page 1)...")
        page1_found = False
        PAGE1_SELS = [
            '[aria-label="Page 1"]',
            '[title="Page 1"]',
            '.reportPageNavigation [role="tab"]:first-child',
            '.pageNavigation button:first-child',
            '[role="tablist"] [role="tab"]:first-child',
            '.navigation-node:first-child',
        ]
        for sel in PAGE1_SELS:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=2_000):
                    btn.click(force=True)
                    page.wait_for_timeout(4_000)
                    self._handle_identity_prompt()
                    page1_found = True
                    log.info(f"  Page 1 clicked via: {sel}")
                    break
            except Exception:
                continue

        if not page1_found:
            log.info("  Could not click Page 1 tab â€” assuming report has only one page, continuing.")

        # One final identity check before touching the slicer
        self._dismiss_popups()
        self._handle_identity_prompt()
        page.wait_for_timeout(2_000)

        # â”€â”€ Step 5: Apply filter via AOM slicer â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log.info(f"  Setting filter: {filter_column} = {filter_email}")
        if not self._try_slicer(filter_email, report_name):
            log.error(
                f"  FILTER APPLY FAILED â€” {aom_name}:\n"
                f"  Could not set '{filter_column}' = '{filter_email}' in slicer.\n"
                f"  This report will NOT be attached."
            )
            self._debug_screenshot(page, f"{aom_name}_{report_name}")
            return None

        # â”€â”€ Step 6: Smart wait â€” poll until other AOMs' data is gone â”€â”€â”€â”€â”€â”€â”€â”€â”€
        MAX_WAIT_SEC  = 90
        POLL_INTERVAL = 5_000
        elapsed_sec   = 0

        log.info(f"  Polling until data refreshes (max {MAX_WAIT_SEC}s)...")
        while elapsed_sec < MAX_WAIT_SEC:
            page.wait_for_timeout(POLL_INTERVAL)
            elapsed_sec += POLL_INTERVAL // 1_000
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(200)
                page_text = page.evaluate("""
                    () => {
                        let clone = document.body.cloneNode(true);
                        clone.querySelectorAll(
                            '.slicer-container, .visual-slicer, ' +
                            '[class*="slicer"], [class*="Slicer"], ' +
                            '.slicerDropdownMenu, [role="listbox"]'
                        ).forEach(el => el.remove());
                        return clone.innerText;
                    }
                """)
                conflicts = [e for e in (other_emails or []) if e in page_text]
                if not conflicts:
                    log.info(f"  Data looks clean after {elapsed_sec}s â€” no conflicting emails. Proceeding.")
                    break
                log.info(f"  [{elapsed_sec}s] Still waiting â€” conflicting emails: {', '.join(conflicts)}")
            except Exception:
                pass

        # -- Step 7: Verify ----------------------------------------------------------------
        verify_ok = self._verify_filter_on_screen(
            filter_email, f"{aom_name}_{report_name}", other_emails or []
        )

        # -- Step 7b: Self-heal if verify failed -----------------------------------------
        # If the report still shows multiple AOMs the slicer may have had more than one
        # item selected. Reopen slicer, deselect all, reselect only correct name, reverify.
        if not verify_ok:
            log.warning("  Verification failed -- self-heal: reopening slicer to fix selection...")
            try:
                healed = self._try_slicer(filter_email, report_name)
                if healed:
                    log.info("  Self-heal slicer done. Re-verifying report...")
                    page.wait_for_timeout(5_000)
                    verify_ok = self._verify_filter_on_screen(
                        filter_email, f"{aom_name}_{report_name}_healed", other_emails or []
                    )
                    if verify_ok:
                        log.info("  Self-heal successful -- data is now clean.")
                    else:
                        log.error("  Self-heal did not resolve the conflict. Aborting export.")
                else:
                    log.error("  Self-heal could not reapply slicer. Aborting export.")
            except Exception as e:
                log.error(f"  Self-heal attempt failed ({e}). Aborting export.")
                verify_ok = False

        if not verify_ok:
            return None


        # â”€â”€ Step 8: Export â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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


    def _try_slicer(self, filter_email: str, report_name: str = "") -> bool:
        """
        Apply the AOM filter via the slicer visual.

        Flow (for each slicer candidate):
          1. Force-unhide all visuals via JS
          2. PRIMARY: Find the visual titled exactly "AOM", open it
          3. Wait 7s for dropdown to render
          4. Deselect "Select all" first (clean slate)
          5. Use the slicer's internal search box (scoped inside dropdown) to type the name
          6. Click the single matching option
          7. Close dropdown
          8. If primary fails â†’ FALLBACK: hunt all slicer triggers
        """
        page = self._page

        # â”€â”€ 1. Force-unhide all slicer visuals â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log.info("  Neutralizing any hidden visibility on visuals via JS...")
        try:
            page.evaluate("""
                () => {
                    document.querySelectorAll('.visual, .visual-container').forEach(el => {
                        el.style.setProperty('display', 'block', 'important');
                        el.style.setProperty('visibility', 'visible', 'important');
                    });
                    document.querySelectorAll('.visual, .visual-container').forEach(el => {
                        let isSlicer = el.querySelector('[role="combobox"]') ||
                                       el.querySelector('.slicerDropdownMenu') ||
                                       el.querySelector('.slicer-container') ||
                                       el.querySelector('.visual-slicer');
                        if (!isSlicer) {
                            el.style.setProperty('pointer-events', 'none', 'important');
                            el.style.setProperty('opacity', '0.1', 'important');
                            el.style.setProperty('z-index', '1', 'important');
                        } else {
                            el.style.setProperty('pointer-events', 'auto', 'important');
                            el.style.setProperty('opacity', '1', 'important');
                            el.style.setProperty('z-index', '2147483647', 'important');
                        }
                    });
                }
            """)
            page.wait_for_timeout(1_000)
        except Exception as e:
            log.warning(f"  JS unhide failed (ignoring): {e}")

        # â”€â”€ Helper: open a slicer trigger and run the select flow â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        def _open_and_select(trigger_el, label: str) -> bool:
            """
            Open the dropdown, deselect all, search for filter_email via the
            slicer's own scoped search box, click the matching option.
            Returns True if the name was successfully selected.
            """
            # Clean single click only — no keyboard shortcuts before clicking.
            # Alt+ArrowDown opens dropdown, then Space toggles item, then click()
            # CLOSES the already-open dropdown (toggle behaviour). Remove all keyboard ops.
            try:
                trigger_el.click(force=True, timeout=1_000)
            except Exception:
                pass

            # Wait 7 seconds for dropdown items to render
            page.wait_for_timeout(7_000)

            # Scope options check to the VISIBLE dropdown container — not page-wide.
            # page.locator('[role="option"]') matches stale DOM from other controls.
            dropdown_open = page.evaluate("""
                () => {
                    const menus = Array.from(document.querySelectorAll(
                        '.slicerDropdownMenu, [role="listbox"], [class*="dropdownMenu"]'
                    )).filter(m => m.offsetHeight > 0 && m.querySelectorAll(
                        '[role="option"], li, span'
                    ).length > 0);
                    return menus.length > 0;
                }
            """)
            if not dropdown_open:
                log.info(f"  {label}: no visible dropdown after 7s. Skipping.")
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                except Exception:
                    pass
                return False

            log.info(f"  {label}: dropdown is open.")

            # Step A: Deselect "Select all" first (clean slate)
            for sel in [
                '[role="option"]:has-text("Select all")',
                '[role="option"]:has-text("(Select all)")',
                '[role="option"]:has-text("All")',
            ]:
                try:
                    opt = page.locator(sel).first
                    if opt.is_visible(timeout=800):
                        opt.click(force=True, timeout=2_000)
                        page.wait_for_timeout(600)
                        log.info("  'Select all' deselected")
                        break
                except Exception:
                    continue

            # Step B: Scoped search box — only inside the visible dropdown.
            # Only set search_typed=True if the name ACTUALLY APPEARS after typing.
            # If it doesn't appear, the input was the wrong element — clear and scroll.
            search_typed = False
            email_sels = [
                f'[role="option"]:has-text("{filter_email}")',
                f'[role="listbox"] li:has-text("{filter_email}")',
                f'div[role="listbox"] span:has-text("{filter_email}")',
            ]
            try:
                s_box = page.evaluate_handle("""
                    () => {
                        const menus = Array.from(document.querySelectorAll(
                            '.slicerDropdownMenu, [role="listbox"], [class*="dropdownMenu"]'
                        )).filter(m => m.offsetHeight > 0);
                        if (menus.length === 0) return null;
                        const menu = menus[menus.length - 1];
                        return menu.querySelector('input[type="text"], input:not([type="hidden"])');
                    }
                """)
                s_input = s_box.as_element() if s_box else None
                if s_input and s_input.is_visible(timeout=500):
                    s_input.fill(filter_email)
                    page.wait_for_timeout(1_000)
                    if any(page.locator(s).count() > 0 for s in email_sels):
                        search_typed = True
                        log.info(f"  Search box used — '{filter_email}' appeared.")
                    else:
                        log.info(f"  Search box typed but '{filter_email}' did not appear. Clearing.")
                        try:
                            s_input.fill("")
                            page.wait_for_timeout(500)
                        except Exception:
                            pass
            except Exception:
                pass

            # Step C: Click the matching option.
            # max_scroll=2 only when search confirmed the name is visible.
            max_scroll = 2 if search_typed else 50
            clicked = False

            for scroll_attempt in range(max_scroll):
                for sel in email_sels:
                    try:
                        if page.locator(sel).count() > 0:
                            opt = page.locator(sel).first
                            try:
                                opt.evaluate("el => el.click()")
                                page.wait_for_timeout(400)
                            except Exception:
                                pass
                            try:
                                opt.click(force=True, timeout=2_000)
                            except Exception:
                                pass
                            page.wait_for_timeout(1_000)
                            log.info(f"  Selected '{filter_email}' ✓")
                            clicked = True
                            break
                    except Exception:
                        continue
                if clicked:
                    break

                if not search_typed:
                    try:
                        page.evaluate("""
                            () => {
                                const menus = Array.from(document.querySelectorAll(
                                    '.slicerDropdownMenu, [role="listbox"]'
                                )).filter(m => m.offsetHeight > 0);
                                if (menus.length === 0) return;
                                const menu = menus[menus.length - 1];
                                const regions = menu.querySelectorAll('.scrollRegion, .scroll-region');
                                const targets = regions.length > 0 ? Array.from(regions) : [menu];
                                targets.forEach(r => {
                                    r.scrollTop += 200;
                                    r.dispatchEvent(new Event('scroll', {bubbles: true}));
                                    r.dispatchEvent(new MouseEvent('scroll', {bubbles: true}));
                                });
                            }
                        """)
                    except Exception:
                        pass
                    page.wait_for_timeout(400)

            if not clicked:
                log.info(f"  {label}: '{filter_email}' not found after search/scroll.")
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                except Exception:
                    pass
                return False

            # Close dropdown
            page.keyboard.press("Escape")
            page.wait_for_timeout(800)
            return True

        # â”€â”€ 2. PRIMARY: Find the visual titled exactly "AOM" â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log.info("  Attempting to locate slicer titled 'AOM' directly...")
        try:
            aom_trigger = page.evaluate_handle("""
                () => {
                    const titleEls = Array.from(document.querySelectorAll(
                        '.visual-title, [class*="visualTitle"], [class*="title"] span, h2, h3, h4, label'
                    ));
                    const titleEl = titleEls.find(el => el.innerText && el.innerText.trim() === 'AOM');
                    if (!titleEl) return null;
                    const container = titleEl.closest(
                        '.visual-container, .visualContainer, [class*="visual"], [class*="Visual"]'
                    );
                    if (!container) return null;
                    // Force-expose the entire ancestor chain
                    let curr = container;
                    while (curr && curr !== document.body) {
                        curr.style.setProperty('visibility', 'visible', 'important');
                        curr.style.setProperty('opacity', '1', 'important');
                        curr.style.setProperty('display', 'block', 'important');
                        curr.style.setProperty('pointer-events', 'auto', 'important');
                        curr.style.setProperty('z-index', '2147483647', 'important');
                        curr = curr.parentElement;
                    }
                    return container.querySelector(
                        '[role="combobox"], [aria-haspopup="listbox"], [aria-haspopup="true"], ' +
                        '.slicerDropdownMenu, [class*="slicerDropdown"], button'
                    );
                }
            """)
            trigger_el = aom_trigger.as_element() if aom_trigger else None
            if trigger_el:
                log.info("  Found slicer titled 'AOM'. Running select flow...")
                if _open_and_select(trigger_el, "AOM slicer (named)"):
                    return True

                # ── Tier 2: Direct DOM click on hidden slicer options ────────────
                # The slicer is hidden so the dropdown won't visually open, but
                # Power BI keeps [role="option"] elements in the DOM.
                # Playwright's :has-text() reads textContent (not innerText) so it
                # finds options even inside display:none containers.
                log.info("  Tier 2: searching for AOM option directly in DOM...")
                tier2_sels = [
                    f'[role="option"]:has-text("{filter_email}")',
                    f'[role="listitem"]:has-text("{filter_email}")',
                    f'li:has-text("{filter_email}")',
                ]
                for t2_sel in tier2_sels:
                    try:
                        if page.locator(t2_sel).count() > 0:
                            opt = page.locator(t2_sel).first
                            try:
                                opt.evaluate("el => el.click()")
                                page.wait_for_timeout(400)
                            except Exception:
                                pass
                            try:
                                opt.click(force=True, timeout=2_000)
                            except Exception:
                                pass
                            page.wait_for_timeout(1_500)
                            log.info(f"  Tier 2: '{filter_email}' clicked via DOM \u2713")
                            return True
                    except Exception:
                        continue
                log.info("  Tier 2: option not found in DOM. Proceeding to general hunt.")

                log.info("  Falling back to general hunt...")
            else:
                log.info("  Slicer titled 'AOM' not found in DOM. Falling back to general hunt...")
        except Exception as e:
            log.info(f"  Named slicer approach failed ({e}). Falling back to general hunt...")
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass

        # â”€â”€ 3. FALLBACK: Hunt all slicer triggers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        SLICER_TRIGGERS = [
            '.visual [role="combobox"]',
            '.visual [aria-haspopup="listbox"]',
            '.visual [aria-haspopup="true"]',
            '.slicerDropdownMenu',
            '.slicer-dropdown-menu',
            '.slicer-dropdown-toggle',
            '[class*="slicerDropdown"]',
        ]
        potential_triggers = []
        for sel in SLICER_TRIGGERS:
            try:
                page.locator(sel).first.wait_for(state="attached", timeout=2_000)
                potential_triggers.extend(page.locator(sel).all())
            except Exception:
                pass

        if not potential_triggers:
            log.warning("  No slicer triggers found in DOM.")
            return False

        cached_idx = self._slicer_cache.get(report_name)
        indices_to_try = list(range(len(potential_triggers)))
        if cached_idx is not None and cached_idx < len(potential_triggers):
            log.info(f"  Prioritizing cached slicer index [{cached_idx}]...")
            indices_to_try.remove(cached_idx)
            indices_to_try.insert(0, cached_idx)
        else:
            log.info(f"  Fallback: hunting {len(potential_triggers)} slicer(s) for '{filter_email}'...")

        for idx in indices_to_try:
            el = potential_triggers[idx]
            try:
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

                if _open_and_select(el, f"slicer[{idx}]"):
                    if cached_idx is None:
                        self._slicer_cache[report_name] = idx
                        log.info(f"  Caching slicer index {idx} for future runs.")
                    return True
                else:
                    if cached_idx is not None and idx == cached_idx:
                        log.warning(f"  Cached slicer [{idx}] failed. Invalidating cache.")
                        self._slicer_cache.pop(report_name, None)
                        cached_idx = None
            except Exception:
                continue

        log.warning("  None of the slicers contained the target name.")
        return False


    def _apply_filter_via_pane(self, filter_email: str, filter_column: str) -> bool:
        """
        Set the AOM Mail Id filter directly through the Power BI Filters pane UI.
        Falls back to Page 1 if the slicer isn't on the current page.
        """
        page = self._page
        original_url = page.url

        # â”€â”€ Primary: interact with the slicer dropdown on the current page â”€â”€â”€â”€
        log.info("  Step A: Opening slicer dropdown on current page...")
        if self._try_slicer(filter_email, ""):
            return True

        # â”€â”€ Fallback 1: Hidden Slicer on Page 1 â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log.info("  Slicer not found on current page. Trying Page 1 fallback...")
        try:
            page.locator('button[aria-label^="Page 1"], .pageNavigation button, .explorationContainer .navigation-node').first.click(timeout=5_000)
            log.info("  Navigated to Page 1.")
            page.wait_for_timeout(5_000)

            if self._try_slicer(filter_email, ""):
                log.info("  Successfully filtered via Page 1 slicer. Navigating back to original page...")
                page.goto(original_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                page.wait_for_timeout(4_000)
                return True
            else:
                log.info("  Slicer not found or failed on Page 1 as well. Navigating back...")
                page.goto(original_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                page.wait_for_timeout(4_000)
        except Exception as e:
            log.warning(f"  Page 1 fallback failed: {e}")

        # â”€â”€ Fallback 2: Filters pane filter card â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        log.info(f"  Step B: Trying Filters pane card for '{filter_column}'...")
        card_found = False
        try:
            # We must search only inside the filter pane, otherwise we might click a table header
            card = page.locator('filter-pane, .filterPane, [aria-label="Filters"]').get_by_text(filter_column, exact=True).first
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

        # Deselect "(Select all)" in the filter card
        for text in ["(Select all)", "Select all"]:
            try:
                el = page.get_by_text(text, exact=True).first
                if el.is_visible(timeout=2_000):
                    el.click(force=True)
                    page.wait_for_timeout(700)
                    log.info(f"  Deselected '{text}' in filter card âœ“")
                    break
            except Exception:
                continue

        # Select the email using role="option" â€” safe, won't hit data table
        for sel in [
            f'[role="option"]:has-text("{filter_email}")',
            f'label:has-text("{filter_email}")',
        ]:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=3_000):
                    el.click(force=True)
                    page.wait_for_timeout(1_500)
                    log.info(f"  Selected '{filter_email}' in filter card âœ“")
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

