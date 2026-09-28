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
import json
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

    def _scoped_verify_text_js(self, slicer_label: str, verify_heading: str = None) -> str:
        """
        Build the JS used to extract "verifiable" page text for the
        conflict/presence checks. Scoped to a "{slicer_label} Wise Sales
        Summary"-style table heading when present, because other tables in
        the report (Cluster-wise, Zone/OPS Head-wise) show staff names
        (Cluster Manager, Operation Head, ...) that can coincidentally match
        a DIFFERENT target's name and cause a false "data conflict" even
        though the actual filter is correct. Falls back to the whole page
        (minus slicer DOM, which stores every option's text even while
        closed) if that table heading can't be found.

        `verify_heading` lets a specific report override the default
        "{slicer_label} Wise Sales Summary" heading pattern if its actual
        table heading text doesn't follow that convention.
        """
        if verify_heading:
            heading_pattern = verify_heading
        else:
            # Escape the label for regex use, then let any whitespace inside
            # it (e.g. "Cluster Manager") match flexibly (\s*) same as the
            # rest of the pattern, rather than requiring an exact single space.
            escaped_label = re.sub(r"\s+", r"\\s*", re.escape(slicer_label))
            heading_pattern = escaped_label + r"\s*Wise\s*Sales\s*Summary"
        heading_js = json.dumps(heading_pattern)
        return r"""
        () => {
            // Clone the body so we don't mutate the live DOM
            let clone = document.body.cloneNode(true);
            // Remove all slicer-related elements — they store all option
            // names in the DOM even when the dropdown is closed, which
            // would cause false "conflict" detections.
            clone.querySelectorAll(
                '.slicer-container, .visual-slicer, ' +
                '[class*="slicer"], [class*="Slicer"], ' +
                '.slicerDropdownMenu, [role="listbox"], ' +
                '[aria-label*="slicer"], [aria-label*="Slicer"]'
            ).forEach(el => el.remove());

            // Try to scope to the target summary table specifically — that's
            // the only section that reflects the actual target-level filter.
            const leafNodes = Array.from(clone.querySelectorAll('*')).filter(
                el => el.children.length === 0
            );
            const headingRe = new RegExp(__HEADING_PATTERN__, 'i');
            const heading = leafNodes.find(
                el => headingRe.test(el.textContent || '')
            );
            if (heading) {
                let container = heading;
                for (let i = 0; i < 6 && container.parentElement; i++) {
                    container = container.parentElement;
                    if (container.querySelectorAll('table, [role="grid"], [role="row"]').length > 0) {
                        break;
                    }
                }
                return container.innerText;
            }
            return clone.innerText;
        }
        """.replace("__HEADING_PATTERN__", heading_js)

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
        self._loaded_report_url = None  # url of the report currently loaded in the page (None = not loaded yet)
        self._loaded_report_state = None  # (url, page, date_from, date_to) tuple currently prepared in the page

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
            if interstitial_email.is_visible(timeout=15_000):
                log.info("  Power BI interstitial detected. Re-submitting email...")
                interstitial_email.fill(self.username)
                
                # Try clicking Submit button if it exists
                submit_btn = page.locator('button:has-text("Submit")').first
                if submit_btn.is_visible(timeout=1_000):
                    submit_btn.click(force=True)
                else:
                    page.keyboard.press("Enter")
                    
                # Wait for the interstitial to disappear
                interstitial_email.wait_for(state="hidden", timeout=15_000)
                page.wait_for_timeout(2_000)
        except Exception as e:
            log.debug(f"  Interstitial handling note: {e}")

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

    def _handle_identity_prompt(self, quick: bool = False) -> None:
        """
        Power BI shows 'You need to verify your identity' dialog at various points.
        Always click 'Continue'.

        quick=False (default): loops up to 5 times with a 3s wait each, to
        robustly catch it appearing/re-appearing after a navigation/reload.
        quick=True: a single, cheap ~300ms check — use this in hot loops
        (e.g. the scroll-fallback loop) where we call this defensively on
        every iteration and can't afford a 3s no-op wait each time.
        """
        max_attempts = 1 if quick else 5
        wait_timeout = 300 if quick else 3_000
        for attempt in range(1, max_attempts + 1):
            try:
                btn = self._page.locator(
                    'button:has-text("Continue"), '
                    '[aria-label="Continue"]'
                ).first
                btn.wait_for(timeout=wait_timeout, state="visible")
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
        self._loaded_report_url = None
        self._loaded_report_state = None
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
        Return the base report URL.

        If `known_url` is given, navigate straight to it -- no workspace
        detour needed once we're already logged in. (The sidebar-based
        workspace lookup below is fragile once the browser is deep inside
        a different report's view -- e.g. after processing 41 targets on a
        prior report, the Filters pane and accumulated UI state mean the
        sidebar workspace link may no longer be reachable the way it was
        right after login. Every report in this codebase's config has a
        known `url`, so this path is what actually runs in practice; the
        sidebar search below only serves as a fallback for a report_cfg
        that omits `url`.)
        """
        page = self._page

        if known_url:
            self._handle_identity_prompt()
            self._dismiss_popups()
            log.info(f'  Navigating to report via known URL...')
            page.goto(known_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(8_000)   # wait for all report visuals to initialise
            self._handle_identity_prompt()
            report_url = page.url.split("?")[0].rstrip("/")
            log.info(f'  Report URL: {report_url}')
            return report_url

        log.info(f'  Looking for workspace "{workspace}"...')

        # Clear any identity prompts or welcome dialogs before looking
        self._handle_identity_prompt()
        self._dismiss_popups()

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

        # (known_url case already returned early above -- this fallback path
        # only runs when report_cfg omits `url` entirely.)
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

    def prepare_report(self, report_cfg: dict, slicer_label: str) -> str:
        """
        Navigate to `report_cfg`'s report, open the Filters pane, set the
        date range and select the target page — but ONLY if this exact
        (url, page, date_from, date_to) combination isn't already the
        currently-prepared state, so calling this once per report per
        group-phase (before looping targets) is cheap on repeat calls.

        Returns the report URL (used by export_for_target for the attempt-3
        reload target and PDF/screenshot naming).
        """
        workspace   = report_cfg["workspace"]
        report_name = report_cfg["name"]
        known_url   = report_cfg.get("url")
        page_name   = report_cfg.get("page")
        date_from   = report_cfg.get("date_from")
        date_to     = report_cfg.get("date_to")

        # Discover report URL (cached after first call per workspace/report).
        if (workspace, report_name) not in self._report_urls:
            self._report_urls[(workspace, report_name)] = self._discover_report_url(
                workspace, report_name, known_url=known_url
            )
        report_url = self._report_urls[(workspace, report_name)]

        page = self._page
        state_key = (report_url, page_name, date_from, date_to)

        if self._loaded_report_state != state_key:
            log.info(f"  Navigating to report (clean state)...")
            page.goto(report_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(4_000)
            # Identity prompt can appear twice in a row right after a reload.
            self._handle_identity_prompt()
            self._handle_identity_prompt()

            log.info(f"  Opening Filters pane...")
            self._open_filters_pane()
            page.wait_for_timeout(1_500)

            if date_from and date_to:
                self._set_date_filter(date_from, date_to)
            if page_name:
                self._select_report_page(page_name)

            self._loaded_report_state = state_key
            self._loaded_report_url = report_url

        return report_url

    def export_for_target(self, target: dict, report_cfg: dict, slicer_label: str,
                           other_targets: list, date_str: str) -> Optional[str]:
        """
        For ONE target within an already-`prepare_report`'d report:
          1. Apply filter via slicer (3-attempt cycle, reload only on the 3rd)
             -> Filters pane card fallback
          2. Smart-wait until conflicting data (from other_targets) disappears
          3. Verify only the expected target's data is on screen
          4. Export as PDF

        Returns local PDF path on success, None on failure.
        """
        workspace      = report_cfg["workspace"]
        report_name    = report_cfg["name"]
        filter_column  = report_cfg.get("filter_column", slicer_label)
        page_name      = report_cfg.get("page")
        date_from      = report_cfg.get("date_from")
        date_to        = report_cfg.get("date_to")
        verify_heading = report_cfg.get("verify_heading")

        report_url = self._report_urls.get((workspace, report_name))
        if not report_url:
            # Defensive: should already be set by prepare_report().
            report_url = self.prepare_report(report_cfg, slicer_label)

        target_name  = target["name"]
        filter_value = target["columns"].get(filter_column)
        if not filter_value:
            log.error(f"  Target '{target_name}' has no value for column '{filter_column}' — skipping.")
            return None

        other_values = [t["columns"].get(filter_column) for t in other_targets]
        other_values = [v for v in other_values if v]

        page = self._page

        def _reprepare_after_reload():
            # Called by _apply_filter_via_pane's attempt-3 reload path to
            # re-apply the page/date state lost by the reload — otherwise
            # subsequent targets in this report would silently export the
            # wrong page/date range.
            if date_from and date_to:
                self._set_date_filter(date_from, date_to)
            if page_name:
                self._select_report_page(page_name)
            self._loaded_report_state = (report_url, page_name, date_from, date_to)

        log.info(f"  Setting filter: {filter_column} = {filter_value}")
        if not self._apply_filter_via_pane(filter_value, filter_column, report_url,
                                            slicer_label, reprepare_fn=_reprepare_after_reload):
            log.error(
                f"  FILTER APPLY FAILED \u2014 {target_name}:\n"
                f"  Could not set '{filter_column}' = '{filter_value}' in UI.\n"
                f"  This report will NOT be attached."
            )
            self._debug_screenshot(page, f"{target_name}_{report_name}")
            return None

        # Smart wait: poll until other targets' data disappears from the page
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
                page_text = page.evaluate(self._scoped_verify_text_js(slicer_label, verify_heading))
                page_text_lower = page_text.lower()
                conflicts = [v for v in other_values if v.lower() in page_text_lower]
                if not conflicts:
                    log.info(
                        f"  Data looks clean after {elapsed_sec}s "
                        f"\u2014 no other {slicer_label} values found. Proceeding."
                    )
                    break
                log.info(
                    f"  [{elapsed_sec}s] Still waiting \u2014 "
                    f"conflicting values present: {', '.join(conflicts)}"
                )
            except Exception:
                pass

        if not self._verify_filter_on_screen(filter_value, f"{target_name}_{report_name}",
                                              other_values, slicer_label, verify_heading):
            return None

        # Export
        safe_target = "".join(c if c.isalnum() or c in " _-" else "_" for c in target_name).strip().replace(" ", "_")
        safe_rep = "".join(c if c.isalnum() or c in " _-" else "_" for c in report_name).strip().replace(" ", "_")
        fname = f"{safe_target}_{safe_rep}_{date_str}.pdf"
        fpath = os.path.join(EXPORTS_DIR, fname)

        try:
            return self._trigger_pdf_export(page, fpath, f"{target_name}_{report_name}",
                                             only_current_page=page_name is not None)
        except Exception as e:
            log.error(f"  Export failed: {e}")
            self._debug_screenshot(page, f"{target_name}_{report_name}")
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

        log.warning("  Could not confirm Filters pane is open — continuing anyway.")
        return False

    def _set_date_filter(self, date_from: str, date_to: str) -> bool:
        """
        Set the Date range slicer's From/To fields. `date_from`/`date_to`
        must already be resolved to the UI's expected display format
        (M/D/YYYY, no leading zeros, e.g. "9/1/2026" -- see main.py's
        resolve_date()).

        NOTE: this is the least-tested part of the automation (no live
        browser access to confirm Power BI's exact date-picker interaction
        during development). Checks the displayed value after each field is
        set and logs clearly either way, so a failure here is diagnosable
        from the log alone without another live-debug round-trip.
        """
        page = self._page
        log.info(f"  Setting Date slicer: {date_from} -> {date_to}...")

        try:
            # Find the "Date" slicer's container (walk up from the label text
            # until we find an ancestor holding 2+ <input> fields).
            container_h = page.evaluate_handle(r"""
            () => {
                const walker = document.createTreeWalker(
                    document.body, NodeFilter.SHOW_TEXT, null
                );
                let node;
                while ((node = walker.nextNode())) {
                    if (node.textContent.trim() !== 'Date') continue;
                    let el = node.parentElement;
                    for (let i = 0; i < 8; i++) {
                        if (!el || el === document.body) break;
                        if (el.querySelectorAll('input').length >= 2) return el;
                        el = el.parentElement;
                    }
                }
                return null;
            }
            """)
            container = container_h.as_element()
            if not container:
                log.warning("  Date slicer container not found — skipping date filter.")
                return False

            date_inputs = container.query_selector_all('input')
            if len(date_inputs) < 2:
                log.warning(f"  Date slicer: expected 2 date inputs, found {len(date_inputs)} — skipping.")
                return False

            from_input, to_input = date_inputs[0], date_inputs[1]

            def _set_one(input_handle, value: str, label: str) -> bool:
                try:
                    input_handle.click(force=True, timeout=2_000)
                    page.wait_for_timeout(200)
                    input_handle.press("Control+A")
                    input_handle.press("Backspace")
                    input_handle.type(value, delay=30)
                    page.wait_for_timeout(200)
                    input_handle.press("Tab")
                    page.wait_for_timeout(500)
                    try:
                        actual = input_handle.input_value()
                    except Exception:
                        actual = None
                    if actual is not None and actual.strip() == value.strip():
                        log.info(f"  Date {label} set to '{value}' ✓")
                        return True
                    log.warning(
                        f"  Date {label}: typed '{value}' but field now shows "
                        f"'{actual}' — may not have applied."
                    )
                    return False
                except Exception as e:
                    log.warning(f"  Date {label} set failed: {e}")
                    return False

            ok_from = _set_one(from_input, date_from, "From")
            ok_to = _set_one(to_input, date_to, "To")
            page.keyboard.press('Escape')  # close any date-picker popup left open
            page.wait_for_timeout(1_000)
            self._handle_identity_prompt(quick=True)
            return ok_from and ok_to
        except Exception as e:
            log.warning(f"  _set_date_filter failed: {e}")
            return False

    def _select_report_page(self, page_name: str) -> bool:
        """
        Click the given page in the left Pages panel to make it the active
        page. Used before exporting when a report_cfg specifies a single
        "page" to export instead of the whole multi-page report.
        """
        page = self._page
        log.info(f"  Selecting page '{page_name}'...")
        try:
            # Page names can be truncated with an ellipsis in the UI, so try
            # an exact match first, then a "starts with" fallback, then the
            # title attribute (often holds the untruncated name as a tooltip).
            candidates = [
                page.get_by_text(page_name, exact=True),
                page.locator(f'[title="{page_name}"]'),
                page.get_by_text(re.compile("^" + re.escape(page_name))),
            ]
            for loc in candidates:
                try:
                    el = loc.first
                    if el.is_visible(timeout=2_000):
                        el.click(force=True, timeout=3_000)
                        page.wait_for_timeout(2_000)
                        log.info(f"  Page '{page_name}' selected.")
                        return True
                except Exception:
                    continue
            log.warning(
                f"  Could not find page '{page_name}' in the Pages panel — "
                f"export may include the wrong page."
            )
            return False
        except Exception as e:
            log.warning(f"  _select_report_page failed: {e}")
            return False

    def _reset_other_slicers(self, skip_label: str) -> None:
        """
        CRITICAL: The report is saved with 'Operation Support = Anuradha Mishra'
        which cross-filters the target slicer to show only a handful of names.
        Reset every slicer EXCEPT `skip_label` (the slicer we're about to set)
        to 'All' so every value appears.

        Uses JavaScript to:
        1. Force-reveal and JS-click all eraser buttons (fast, works off-screen)
        2. Falls back to Playwright hover+click if JS found nothing
        """
        page = self._page
        log.info(f"  Resetting other slicers (clearing cross-filters, keeping '{skip_label}')...")

        # --- Primary: pure JavaScript approach ---
        # Walk the DOM to find the target slicer's visual container, then click
        # every OTHER slicer's eraser button using JS (works even when off-screen).
        n_cleared = page.evaluate(r"""
        (SKIP_LABEL) => {
            let count = 0;

            // ── Step 1: identify the target slicer's direct visual container ──
            // Look for a container whose DIRECT slicer-header text matches
            // (max 6 levels up from the dropdown trigger, not the full page).
            function findTargetContainer() {
                const menus = document.querySelectorAll(
                    '.slicerDropdownMenu, [aria-haspopup="listbox"], [role="combobox"]'
                );
                for (const menu of menus) {
                    let p = menu.parentElement;
                    for (let i = 0; i < 6; i++) {
                        if (!p || p === document.body) break;
                        // Check header-label elements at THIS level only
                        // (:scope selects only direct children of p)
                        const hdrs = p.querySelectorAll(
                            '[class*="slicerHeader"], [class*="headerLabel"], '
                            + '[class*="labelText"], .title, .slicerTitle'
                        );
                        for (const h of hdrs) {
                            // textContent.trim() must EXACTLY match, not just contain
                            if (h.textContent.trim() === SKIP_LABEL) return p;
                        }
                        p = p.parentElement;
                    }
                }
                return null;
            }

            const aomContainer = findTargetContainer();

            // ── Step 2: collect ALL slicer dropdown triggers ──
            const allMenus = document.querySelectorAll(
                '.slicerDropdownMenu, [aria-haspopup="listbox"], [role="combobox"]'
            );

            allMenus.forEach(menu => {
                // Skip if inside the AOM container
                if (aomContainer && aomContainer.contains(menu)) return;

                // Scroll into view so hover events work
                menu.scrollIntoView({ block: 'center', inline: 'center' });

                // Dispatch hover events to cause eraser to appear in DOM
                ['mouseenter', 'mouseover', 'pointermove', 'pointerenter'].forEach(type => {
                    try {
                        menu.dispatchEvent(
                            new MouseEvent(type, { bubbles: true, cancelable: true, view: window })
                        );
                    } catch(ex) {}
                });

                // Walk up to find the eraser button
                let p = menu.parentElement;
                let found = false;
                for (let i = 0; i < 10 && !found; i++) {
                    if (!p || p === document.body) break;

                    // Also dispatch hover on parent to trigger CSS :hover
                    try {
                        p.dispatchEvent(
                            new MouseEvent('mouseover', { bubbles: true, view: window })
                        );
                    } catch(ex) {}

                    const ERASER_SELS = [
                        '.slicerDeleteButton',
                        '[class*="clearButton"]',
                        '[class*="eraserButton"]',
                        '[class*="slicerDelete"]',
                        '[aria-label*="clear" i]',
                        '[aria-label*="eraser" i]',
                        '[title*="clear" i]',
                        '[title*="Remove filter" i]',
                        '[title*="Remove" i]',
                    ];

                    for (const sel of ERASER_SELS) {
                        const eraser = p.querySelector(sel);
                        if (eraser) {
                            // Force visible
                            eraser.style.cssText +=
                                ';display:block!important'
                                + ';visibility:visible!important'
                                + ';opacity:1!important'
                                + ';pointer-events:auto!important';
                            eraser.click();
                            count++;
                            found = true;
                            break;
                        }
                    }
                    p = p.parentElement;
                }
            });

            return count;
        }
        """, skip_label)

        log.info(f"  JavaScript cleared {n_cleared} other slicer(s).")

        # --- Fallback: Playwright hover approach (for elements JS couldn't clear) ---
        if n_cleared == 0:
            log.warning(
                "  JS cleared 0 slicers — falling back to Playwright hover approach.")
            try:
                all_menus = page.locator(
                    '.slicerDropdownMenu, [aria-haspopup="listbox"], [role="combobox"]'
                ).all()
                log.info(f"  Playwright found {len(all_menus)} slicer triggers.")
                for i, menu in enumerate(all_menus):
                    try:
                        # is_target: only look 5 levels up and check specific header classes
                        is_target = menu.evaluate("""
                        (el, SKIP_LABEL) => {
                            let p = el.parentElement;
                            for (let i = 0; i < 5; i++) {
                                if (!p || p === document.body) break;
                                const hdrs = p.querySelectorAll(
                                    '[class*="slicerHeader"], [class*="headerLabel"], '
                                    + '[class*="labelText"], .title, .slicerTitle'
                                );
                                for (const h of hdrs) {
                                    if (h.textContent.trim() === SKIP_LABEL) return true;
                                }
                                p = p.parentElement;
                            }
                            return false;
                        }
                        """, skip_label)
                        if is_target:
                            log.info(f"  Slicer {i}: '{skip_label}' — skipping.")
                            continue

                        # Scroll into view
                        page.evaluate(
                            "(el) => el.scrollIntoView({block:'center',inline:'center'})",
                            menu)
                        page.wait_for_timeout(300)

                        # Hover
                        try:
                            menu.hover(timeout=2_000)
                            page.wait_for_timeout(400)
                        except Exception:
                            pass

                        # Find eraser
                        eraser_h = page.evaluate_handle("""
                        (el) => {
                            let p = el.parentElement;
                            for (let i = 0; i < 8; i++) {
                                if (!p || p === document.body) break;
                                const e = p.querySelector(
                                    '.slicerDeleteButton, [class*="clearButton"], '
                                    + '[aria-label*="clear" i], [title*="clear" i]'
                                );
                                if (e) {
                                    e.style.cssText += ';display:block!important'
                                        + ';visibility:visible!important'
                                        + ';opacity:1!important'
                                        + ';pointer-events:auto!important';
                                    return e;
                                }
                                p = p.parentElement;
                            }
                            return null;
                        }
                        """, menu)
                        eraser = eraser_h.as_element()
                        if eraser:
                            eraser.click(force=True, timeout=1_500)
                            page.wait_for_timeout(300)
                            log.info(f"  Slicer {i}: cleared via hover.")
                        else:
                            val = ''
                            try:
                                val = menu.text_content()[:30]
                            except Exception:
                                pass
                            log.info(f"  Slicer {i}: no eraser found (val='{val}').")
                    except Exception as ex:
                        log.info(f"  Slicer {i}: {type(ex).__name__}")
                        continue
            except Exception as ex:
                log.warning(f"  Fallback hover approach failed: {ex}")

        page.wait_for_timeout(2_000)  # Let report re-render after clearing

    def _try_slicer(self, filter_email: str, slicer_label: str = "AOM") -> bool:
        """
        Attempts to select the given value in the `slicer_label` slicer's
        dropdown, ONE time. Implements a Search + Scroll Fallback strategy
        within this single attempt:
        1. Types the name in the internal search box.
        2. If not found, scrolls down the list sequentially looking for it,
           stopping early once the visible names have passed the target
           alphabetically (the list is sorted A-Z).
        Returns True on success, False on failure. Retrying (in place, or via
        a page reload) is the caller's responsibility (_apply_filter_via_pane).
        """
        page = self._page
        slicer_label_js = json.dumps(slicer_label)

        # JS: find the target slicer trigger (prefer .slicerDropdownMenu)
        FIND_TRIGGER_JS = r"""
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let node;
            while ((node = walker.nextNode())) {
                if (node.textContent.trim() !== __LABEL__) continue;
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
        """.replace("__LABEL__", slicer_label_js)

        # JS: find eraser/clear button near the target slicer's label text
        FIND_ERASER_JS = r"""
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let node;
            while ((node = walker.nextNode())) {
                if (node.textContent.trim() !== __LABEL__) continue;
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
        """.replace("__LABEL__", slicer_label_js)

        OPTION_SELS = '[role="option"], [role="listbox"] li, div[role="listbox"] span'
        EMAIL_SELS = [
            f'[role="option"]:has-text("{filter_email}")',
            f'[role="option"][title="{filter_email}" i]',
            f'[role="option"][aria-label="{filter_email}" i]'
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

        def _is_dropdown_open() -> bool:
            try:
                return page.locator(OPTION_SELS).count() > 0
            except Exception:
                return False

        def _open_dropdown(elem_h):
            """
            Open the AOM dropdown. IMPORTANT: this trigger toggles open/closed,
            so we try ONE action at a time and check whether options actually
            appeared before trying the next — firing multiple toggle actions
            back-to-back (click, then Enter, then Space, ...) risks opening
            and immediately re-closing it, which was causing intermittent
            failures where the dropdown silently ended up closed.
            """
            _unhide_chain(elem_h)
            page.wait_for_timeout(200)

            # Attempt 1: plain click.
            try:
                elem_h.click(force=True, timeout=2_000)
            except Exception:
                pass
            page.wait_for_timeout(500)
            if _is_dropdown_open():
                return

            # Attempt 2: focus + Enter (only if click didn't open it).
            try:
                elem_h.focus()
                page.wait_for_timeout(150)
                page.keyboard.press('Enter')
                page.wait_for_timeout(500)
            except Exception:
                pass
            if _is_dropdown_open():
                return

            # Attempt 3: Space (last resort — still only fired if still closed).
            try:
                page.keyboard.press('Space')
                page.wait_for_timeout(500)
            except Exception:
                pass

        def _check_target_visible() -> bool:
            return any(page.locator(s).count() > 0 for s in EMAIL_SELS)

        def _get_top_visible_option_text() -> str:
            """First (topmost) currently-rendered option's text, used for the
            alphabetical scroll early-stop (the list is sorted A-Z)."""
            try:
                return page.evaluate(f"""
                () => {{
                    const els = document.querySelectorAll('{OPTION_SELS}');
                    for (const e of els) {{
                        const t = (e.innerText || e.textContent || '').trim();
                        if (t) return t;
                    }}
                    return '';
                }}
                """) or ''
            except Exception:
                return ''

        def _click_target() -> bool:
            # Single click only \u2014 Power BI slicer options are checkbox-style;
            # a second click risks toggling the just-made selection back off.
            for sel in EMAIL_SELS:
                try:
                    loc = page.locator(sel)
                    if loc.count() > 0:
                        opt = loc.first
                        self._handle_identity_prompt()
                        opt.evaluate('el => el.click()')   # JS click (bypasses overlay)
                        page.wait_for_timeout(1_500)
                        log.info(f'  Selected \'{filter_email}\' in {slicer_label} slicer \u2713')
                        page.keyboard.press('Escape')
                        page.wait_for_timeout(600)
                        return True
                except Exception:
                    pass
            return False

        log.info(f"  {slicer_label} slicer: attempt for '{filter_email}'...")

        self._handle_identity_prompt()
        self._dismiss_popups()

        # Find the slicer trigger (the dropdown caret)
        try:
            h = page.evaluate_handle(FIND_TRIGGER_JS)
            elem = h.as_element()
        except Exception:
            elem = None

        if not elem:
            log.warning(f"  {slicer_label} slicer trigger not found in DOM.")
            return False

        # Clear eraser if a previous selection is visible
        try:
            elem.hover(timeout=1_000)
            page.wait_for_timeout(300)
            eh = page.evaluate_handle(FIND_ERASER_JS)
            eraser = eh.as_element()
            if eraser:
                _unhide_chain(eraser)
                eraser.click(force=True, timeout=1_000)
                page.wait_for_timeout(800)
                log.info('  Eraser clicked — previous selection cleared.')
        except Exception:
            pass

        # Open dropdown
        _open_dropdown(elem)

        try:
            page.wait_for_selector(OPTION_SELS, timeout=8_000, state='attached')
        except Exception:
            log.info("  No dropdown options appeared.")
            page.keyboard.press('Escape')
            page.wait_for_timeout(1_000)
            return False

        self._handle_identity_prompt()

        # ── 1. THE SEARCH STEP ──
        search_typed = False
        try:
            # Target the search input specifically inside the opened dropdown container
            sb_h = page.evaluate_handle("""
            () => {
                const opt = document.querySelector('[role="option"]');
                if (!opt) return null;
                let p = opt.parentElement;
                for (let i = 0; i < 10; i++) {
                    if (!p || p === document.body) break;
                    const inp = p.querySelector('input');
                    if (inp) return inp;
                    p = p.parentElement;
                }
                return null;
            }
            """)
            sb = sb_h.as_element()
            if sb:
                sb.click(force=True, timeout=1_000)
                page.wait_for_timeout(200)
                sb.fill(filter_email)  # Force fill value
                page.wait_for_timeout(200)
                sb.type(' ', delay=10) # Trigger change event
                page.keyboard.press('Backspace')
                # DO NOT press 'Enter' here — it submits/closes the dropdown in Power BI!
                search_typed = True
                log.info(f"  Search box: typed '{filter_email}'")
            else:
                n_opts = page.locator(OPTION_SELS).count()
                log.warning(
                    f"  Search box not found in dropdown (options currently visible: {n_opts})."
                )
        except Exception as e:
            log.warning(f"  Search box fill failed: {e}")

        if search_typed:
            page.wait_for_timeout(4_000)  # Wait for virtual list to update (increased wait)

        if _check_target_visible():
            if _click_target():
                return True
            log.warning("  Target was visible but click failed.")
            page.keyboard.press('Escape')
            page.wait_for_timeout(1_000)
            return False

        log.info(f"  Target '{filter_email}' not immediately visible. Falling back to SCROLLING...")

        # ── 2. THE SCROLL FALLBACK STEP ──
        # If the target is further down the virtualized list, we scroll to find it.
        # We click inside the dropdown list container and press PageDown, stopping
        # early once the visible names have passed the target alphabetically
        # (the list is sorted A-Z, so there's no point scrolling further).
        scroll_success = False
        target_letter = filter_email.strip()[0].lower() if filter_email.strip() else ''
        try:
            # Focus the list container so keystrokes scroll it
            page.evaluate("""
            () => {
                const opt = document.querySelector('[role="option"]');
                if (opt) {
                    // Focus the scrollable viewport WITHOUT clicking any option
                    // (clicking an option here would accidentally select the
                    // wrong AOM — we only want keyboard focus for PageDown).
                    let viewport = opt.closest('.cdk-virtual-scroll-viewport, .scrollable-content, [role="listbox"]');
                    if (viewport && viewport.focus) {
                        viewport.setAttribute('tabindex', viewport.getAttribute('tabindex') || '0');
                        viewport.focus();
                    } else if (opt.focus) {
                        opt.focus(); // focus only, never click
                    }
                }
            }
            """)

            max_scrolls = 15
            for s in range(max_scrolls):
                if _check_target_visible():
                    log.info(f"  Found '{filter_email}' after {s} scrolls!")
                    scroll_success = True
                    break

                # Alphabetical early-stop: if the topmost visible name has
                # already moved past the target's first letter, stop scrolling.
                top_text = _get_top_visible_option_text()
                if target_letter and top_text:
                    top_letter = top_text.strip()[0].lower()
                    if top_letter > target_letter:
                        log.info(
                            f"  Visible list has passed '{filter_email}' alphabetically "
                            f"(now at '{top_text}') — stopping scroll early."
                        )
                        break

                # Press PageDown to scroll the virtual list
                page.keyboard.press('PageDown')
                page.wait_for_timeout(600)  # wait for new items to render
                self._handle_identity_prompt(quick=True) # just in case, cheap check

            if not scroll_success:
                # Diagnostics: dump what was actually visible when we gave up,
                # so a future failure is debuggable from the log instead of
                # requiring another guess-and-check round.
                try:
                    sample = page.evaluate(f"""
                    () => Array.from(document.querySelectorAll('{OPTION_SELS}'))
                        .slice(0, 5)
                        .map(e => (e.innerText || e.textContent || '').trim())
                    """)
                except Exception:
                    sample = []
                log.warning(
                    f"  Could not find '{filter_email}' even after scrolling. "
                    f"Currently visible (sample): {sample}"
                )
        except Exception as e:
            log.warning(f"  Scroll fallback failed: {e}")

        if scroll_success and _click_target():
            return True

        page.keyboard.press('Escape')
        page.wait_for_timeout(1_000)
        return False

    def _apply_filter_via_pane(self, filter_email: str, filter_column: str, report_url: str = None,
                                slicer_label: str = "AOM", reprepare_fn=None) -> bool:
        """
        Set the `slicer_label` filter through the Power BI UI (slicer on the
        current page only). Assumes the slicer is on the first/current page.

        Runs the full 3-attempt cycle:
          - Attempt 1: reset other slicers, try the slicer (search + scroll fallback).
          - Attempt 2: same, in place, no reload.
          - Attempt 3: only if 1 & 2 both failed — reload the report page once,
            handle the identity prompt, re-open the Filters pane, re-apply the
            page/date-range state via `reprepare_fn` (if given), reset other
            slicers again, then try the slicer one last time.
        Falls back to the Filters pane card if all 3 slicer attempts fail.
        """
        page = self._page

        # Attempts 1 & 2: in place, no reload.
        for attempt in (1, 2):
            log.info(f"  {slicer_label} slicer: attempt {attempt}/3 for '{filter_email}'...")
            self._reset_other_slicers(slicer_label)
            if self._try_slicer(filter_email, slicer_label):
                return True

        # Attempt 3: hard reset (reload) then one more try.
        if report_url:
            log.warning(f"  {slicer_label} slicer: attempts 1 & 2 failed. Reloading report for attempt 3/3...")
            page.wait_for_timeout(2_000)
            page.goto(report_url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
            page.wait_for_timeout(5_000)
            self._handle_identity_prompt()
            self._open_filters_pane()
            page.wait_for_timeout(1_500)
            if reprepare_fn is not None:
                # Re-apply page selection + date range lost by the reload —
                # otherwise remaining targets in this report would silently
                # export the wrong page/date after a mid-loop reload.
                try:
                    reprepare_fn()
                except Exception as e:
                    log.warning(f"  reprepare_fn failed after reload: {e}")
            self._reset_other_slicers(slicer_label)
            log.info(f"  {slicer_label} slicer: attempt 3/3 for '{filter_email}' (after reload)...")
            if self._try_slicer(filter_email, slicer_label):
                return True
        else:
            log.warning("  No report_url available for attempt 3 reload — skipping to Filters pane fallback.")

        # Fallback: Filters pane card (search by the slicer's UI label, not
        # necessarily filter_column -- e.g. filter_column may be "AOM Mail Id"
        # while the visible card is still labeled "AOM").
        log.info(f"  Step B: Trying Filters pane card for '{slicer_label}'...")
        card_found = False
        try:
            card = page.locator(
                'filter-pane, .filterPane, [aria-label="Filters"]'
            ).get_by_text(slicer_label, exact=True).first
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
                                  other_emails: list, slicer_label: str = "AOM",
                                  verify_heading: str = None) -> bool:
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
            page_text = page.evaluate(self._scoped_verify_text_js(slicer_label, verify_heading))
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

    def _trigger_pdf_export(self, page, fpath: str, aom_name: str, only_current_page: bool = False) -> str:
        """
        Drive the Power BI UI to export the report as PDF.

        Power BI sometimes downloads the PDF directly (Playwright download event)
        and sometimes opens it in a new browser tab.  This method handles both.

        Flow:
          dismiss popups → click Export (toolbar) → click PDF option →
          confirm dialog (if shown; check "Only export current page" first
          when only_current_page=True) → capture via download event OR new tab.
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

                    # When exporting a single page, check "Only export
                    # current page" in the Export dialog before confirming
                    # (unchecked by default -- leaving it unchecked would
                    # export every page instead of just the selected one).
                    if only_current_page:
                        CURRENT_PAGE_CHECKBOX_SELS = [
                            'text="Only export current page"',
                            '[aria-label="Only export current page"]',
                        ]
                        checked = False
                        for sel in CURRENT_PAGE_CHECKBOX_SELS:
                            try:
                                el = page.locator(sel).first
                                el.wait_for(timeout=3_000, state='visible')
                                el.click(force=True)
                                log.info('  "Only export current page" checked.')
                                checked = True
                                break
                            except Exception:
                                continue
                        if not checked:
                            log.warning(
                                '  Could not find/check "Only export current page" — '
                                'export may include every page instead of just the target one.'
                            )
                        page.wait_for_timeout(500)

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

