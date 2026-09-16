"""
powerbi.py -- Headless Power BI automation via Playwright.

Responsibilities:
  1. Log into Power BI with provided credentials (handles MFA OTP via terminal)
  2. Navigate to the configured workspace and report (no hardcoded URLs)
  3. For each AOM, apply a URL filter and export all pages as PDF

Runs headless (invisible) - no browser window appears during execution.
"""
import os
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
        page.goto("https://login.microsoftonline.com/", timeout=NAV_TIMEOUT, wait_until="networkidle")
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

        Power BI URL filter spec requires:
          - '@' in email values encoded as '%40'
          - Spaces in table/column NAMES encoded as '%20'
          (Without %20, PBI silently ignores the filter and shows all data)
        """
        encoded_value = filter_value.replace("@", "%40")
        table  = PBI_FILTER_TABLE.replace(" ", "%20")
        column = PBI_FILTER_COLUMN.replace(" ", "%20")
        expr   = f"{table}/{column} eq '{encoded_value}'"
        filter_url = f"{self._report_url}?filter={expr}"
        log.info(f"  Filter URL: ...?filter={table}/{column} eq '{encoded_value}'")
        return filter_url

    def export_report(self, filter_email: str, aom_name: str, date_str: str) -> Optional[str]:
        """
        Navigate to the report filtered for this AOM, export all pages as PDF.
        Returns local path of the downloaded PDF, or None on failure.
        """
        page = self._page
        url  = self._build_filter_url(filter_email)

        log.info(f"  Loading filtered report...")
        page.goto(url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
        page.wait_for_timeout(REPORT_WAIT_MS)   # let all visuals render

        # Dismiss "You need to verify your identity" dialog if it appears (up to 2x).
        # IMPORTANT: after clicking Continue, PBI re-authenticates and may redirect
        # the page, losing the ?filter= query parameter. So after dismissing,
        # we always re-navigate to the filtered URL to guarantee the filter is active.
        self._handle_identity_prompt()

        if "filter=" not in page.url:
            log.info("  Filter lost after identity prompt — re-navigating to filtered URL...")
            page.goto(url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(REPORT_WAIT_MS)
        else:
            # Filter is still in URL — give PBI a bit more time to apply it
            page.wait_for_timeout(3_000)

        log.info(f"  Active URL confirms filter: {'yes' if 'filter=' in page.url else 'NO - filter missing!'}")

        # Build output file path
        safe  = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name)
        fname = f"{safe.replace(' ', '_')}_{date_str}.pdf"
        fpath = os.path.join(EXPORTS_DIR, fname)

        try:
            return self._trigger_pdf_export(page, fpath, aom_name)
        except Exception as e:
            log.error(f"  Export failed: {e}")
            self._debug_screenshot(page, aom_name)
            return None


    def _trigger_pdf_export(self, page, fpath: str, aom_name: str) -> str:
        """
        Drive the Power BI UI to export the report as PDF.

        The Export button is a standalone button in the toolbar
        (confirmed from screenshots) — NOT inside the File menu.
        File menu only has: Download this file / Print / Embed / QR code.

        Flow: dismiss popups → click Export (toolbar) → click PDF → download
        """
        # Step 1: Dismiss Copilot dialog, CDK overlays, any popups
        self._dismiss_popups()
        page.wait_for_timeout(500)

        # Step 2: Click the Export toolbar button directly
        # It sits in the top command bar alongside File, Share, Explore, etc.
        log.info("  Clicking Export toolbar button...")
        page.locator(
            '[aria-label="Export"], '
            'button:has-text("Export")'
        ).first.click(force=True, timeout=10_000)
        page.wait_for_timeout(800)

        # Step 3: Select PDF from the dropdown that appears
        log.info("  Selecting PDF...")
        page.locator(
            '[role="menuitem"]:has-text("PDF"), '
            '[role="option"]:has-text("PDF"), '
            'button:has-text("PDF"), '
            'a:has-text("PDF")'
        ).first.click(force=True, timeout=8_000)
        page.wait_for_timeout(1_000)

        # Step 4: If an export options dialog appears, confirm it
        # (Some PBI setups show "Export" confirmation, others start download immediately)
        try:
            confirm_btn = page.locator(
                'button:has-text("Export")'
            ).last
            confirm_btn.wait_for(timeout=5_000, state="visible")
            with page.expect_download(timeout=EXPORT_TIMEOUT) as dl_info:
                confirm_btn.click(force=True)
        except PWTimeout:
            # No confirmation dialog — download already started
            with page.expect_download(timeout=EXPORT_TIMEOUT) as dl_info:
                pass   # download was already triggered by the PDF click

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
