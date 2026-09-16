"""
powerbi.py - Power BI report automation via REST API with effectiveIdentity.

AUTHENTICATION : Playwright headless browser (handles MFA / SSO).
                 After login the Bearer token is captured from intercepted API requests.
EXPORT         : Power BI REST Export API  with  effectiveIdentity.
                 The server applies RLS for the given AOM email address, so the PDF
                 contains ONLY that AOM's stores — no admin bypass, no URL-filter trick.

Flow
----
1. Open a headless Chromium, log in as admin (MFA handled in terminal if needed).
2. Navigate to Power BI home; intercept Bearer token from requests to api.powerbi.com.
3. Close the browser.
4. Fetch the Dataset ID linked to the report (one REST call).
5. For each AOM:
     POST ExportTo with effectiveIdentity {username, roles, datasets}
     Poll until status == Succeeded
     Download the PDF
"""
import os
import time
import logging
import requests
from typing import Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from config import (
    PBI_GROUP_ID, PBI_REPORT_ID, PBI_RLS_ROLE,
    EXPORTS_DIR,
)

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
NAV_TIMEOUT    = 60_000   # ms  — browser navigation
TOKEN_WAIT_S   = 40       # sec — wait for Bearer token
POLL_INTERVAL  = 6        # sec — between export-status polls
EXPORT_TIMEOUT = 300      # sec — give up if export takes too long
PBI_API        = "https://api.powerbi.com/v1.0/myorg"


class PowerBIClient:
    """
    Context manager.

    __enter__  logs in via headless browser, grabs the Bearer token,
               closes the browser, then fetches the dataset ID.
    __exit__   nothing to clean up (token is in memory only).
    """

    def __init__(self, email: str, password: str):
        self.email       = email
        self.password    = password
        self._token:      Optional[str] = None
        self._dataset_id: Optional[str] = None

    # ------------------------------------------------------------------ setup

    def __enter__(self) -> "PowerBIClient":
        log.info("Authenticating with Power BI (headless browser)...")
        self._token = self._login_and_capture_token()
        log.info("Bearer token acquired. Browser closed.")
        self._dataset_id = self._fetch_dataset_id()
        log.info(f"Dataset ID confirmed: {self._dataset_id}")
        return self

    def __exit__(self, *args) -> None:
        pass   # token lives only in memory — nothing to close

    # ------------------------------------------------------------------- auth

    def _login_and_capture_token(self) -> str:
        """
        Log in via headless Chromium and intercept the Bearer token that
        Power BI sends to api.powerbi.com when the home page loads.
        """
        captured = {"token": None}

        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )
            ctx = browser.new_context(
                viewport={"width": 1920, "height": 1080},
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/127.0.0.0 Safari/537.36"
                ),
            )
            page = ctx.new_page()

            # Intercept every outgoing request and grab the first Bearer token
            # that goes to api.powerbi.com
            def on_request(request):
                if captured["token"]:
                    return
                if "api.powerbi.com" not in request.url:
                    return
                auth = request.headers.get("authorization", "")
                if auth.startswith("Bearer "):
                    captured["token"] = auth[7:]
                    log.info("  Bearer token captured from API request.")

            page.on("request", on_request)

            # ----- login sequence (same as previous working version) ------
            self._do_login(page)

            # PBI home triggers multiple api.powerbi.com calls — token appears here
            log.info("  Navigating to Power BI home to capture token...")
            page.goto(
                "https://app.powerbi.com/",
                timeout=NAV_TIMEOUT,
                wait_until="domcontentloaded",
            )
            deadline = time.time() + TOKEN_WAIT_S
            while not captured["token"] and time.time() < deadline:
                page.wait_for_timeout(1_000)

            # Fallback: go directly to the report to force an API call
            if not captured["token"]:
                log.info("  Token not yet seen — navigating to report...")
                page.goto(
                    f"https://app.powerbi.com/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}",
                    timeout=NAV_TIMEOUT,
                    wait_until="domcontentloaded",
                )
                deadline = time.time() + TOKEN_WAIT_S
                while not captured["token"] and time.time() < deadline:
                    page.wait_for_timeout(1_000)

            browser.close()

        if not captured["token"]:
            raise RuntimeError(
                "Bearer token was not captured. "
                "Power BI did not make any API requests within the expected window."
            )
        return captured["token"]

    def _do_login(self, page) -> None:
        """Navigate Microsoft login, fill credentials, handle MFA, skip Stay-signed-in."""
        log.info("  Navigating to Microsoft login page...")
        page.goto(
            "https://login.microsoftonline.com/",
            timeout=NAV_TIMEOUT,
            wait_until="domcontentloaded",
        )
        page.wait_for_timeout(1_500)

        # Step 1 — email
        for sel in ['input[name="loginfmt"]', 'input[type="email"]', "#i0116"]:
            try:
                page.wait_for_selector(sel, timeout=10_000)
                page.fill(sel, self.email)
                page.keyboard.press("Enter")
                log.info("  Email submitted.")
                break
            except PWTimeout:
                continue
        page.wait_for_timeout(2_000)

        # Step 2 — password
        for sel in ['input[name="passwd"]', 'input[type="password"]', "#i0118"]:
            try:
                page.wait_for_selector(sel, timeout=10_000)
                page.fill(sel, self.password)
                page.keyboard.press("Enter")
                log.info("  Password submitted. Checking for MFA...")
                break
            except PWTimeout:
                continue
        page.wait_for_timeout(2_000)

        # Step 3 — MFA (if any)
        self._handle_mfa(page)

        # Step 4 — "Stay signed in?" → always No
        try:
            page.wait_for_selector("#idBtn_Back", timeout=8_000)
            page.click("#idBtn_Back")
            log.info("  'Stay signed in?' -- answered No.")
        except PWTimeout:
            pass

        log.info("  Login complete.")

    def _handle_mfa(self, page) -> None:
        """Handle TOTP or authenticator-app push MFA if it appears."""
        # TOTP / SMS code input
        try:
            page.wait_for_selector(
                'input[name="otc"], #idTxtBx_SAOTCC_OTC', timeout=8_000
            )
            otp = input("\n  MFA code required — enter it here: ").strip()
            page.fill('input[name="otc"], #idTxtBx_SAOTCC_OTC', otp)
            page.keyboard.press("Enter")
            page.wait_for_timeout(3_000)
            log.info("  MFA code submitted.")
            return
        except PWTimeout:
            pass

        # Authenticator app push notification
        try:
            page.wait_for_selector(
                '[data-value="PhoneAppNotification"], #idDiv_SAOTCS_Section',
                timeout=5_000,
            )
            log.info("  Authenticator push sent — approve on your phone (waiting 20 s)...")
            page.wait_for_timeout(20_000)
            return
        except PWTimeout:
            pass

        log.info("  No MFA challenge detected.")

    # --------------------------------------------------------------- REST API

    def _api_headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type":  "application/json",
        }

    def _fetch_dataset_id(self) -> str:
        """GET report metadata to extract the dataset ID linked to this report."""
        url  = f"{PBI_API}/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}"
        resp = requests.get(url, headers=self._api_headers(), timeout=30)
        if resp.status_code == 401:
            raise RuntimeError(
                "Bearer token rejected (401 Unauthorized). "
                "The browser session may not have completed before the token was read."
            )
        resp.raise_for_status()
        return resp.json()["datasetId"]

    # ----------------------------------------------------------- per-AOM export

    def export_report(
        self, filter_email: str, aom_name: str, date_str: str
    ) -> Optional[str]:
        """
        Export the report for one AOM via the REST API with effectiveIdentity.

        filter_email  the AOM's Power BI UPN (e.g. aomnorth1@kisna.com).
                      Passed as the effectiveIdentity username so the server
                      applies RLS and includes only that AOM's stores.
        aom_name      used for the PDF filename only.
        date_str      used for the PDF filename only.

        Returns the local path to the saved PDF, or None on failure.
        """
        os.makedirs(EXPORTS_DIR, exist_ok=True)
        safe  = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name)
        fname = f"{safe.replace(' ', '_')}_{date_str}.pdf"
        fpath = os.path.join(EXPORTS_DIR, fname)

        try:
            log.info(f"  Requesting export (effectiveIdentity: {filter_email})...")
            export_id = self._start_export(filter_email)

            log.info("  Waiting for Power BI to generate the PDF...")
            self._poll_export(export_id)

            log.info("  Downloading PDF...")
            self._download(export_id, fpath)

            log.info(f"  PDF saved: {fname}")
            return fpath

        except Exception as exc:
            log.error(f"  Export failed: {exc}")
            return None

    def _start_export(self, filter_email: str) -> str:
        """POST to ExportTo with effectiveIdentity; return the export job ID."""
        url  = f"{PBI_API}/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}/ExportTo"
        body = {
            "format": "PDF",
            "powerBIReportConfiguration": {
                "identities": [
                    {
                        "username": filter_email,
                        "roles":    [PBI_RLS_ROLE],
                        "datasets": [self._dataset_id],
                    }
                ]
            },
        }
        resp = requests.post(url, headers=self._api_headers(), json=body, timeout=30)
        if resp.status_code not in (200, 202):
            raise RuntimeError(
                f"ExportTo API failed: HTTP {resp.status_code}\n{resp.text[:500]}"
            )
        return resp.json()["id"]

    def _poll_export(self, export_id: str) -> None:
        """Poll export status until Succeeded, Failed, or timeout."""
        url      = f"{PBI_API}/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}/exports/{export_id}"
        deadline = time.time() + EXPORT_TIMEOUT

        while time.time() < deadline:
            resp = requests.get(url, headers=self._api_headers(), timeout=30)
            resp.raise_for_status()
            data   = resp.json()
            status = data.get("status", "Unknown")
            pct    = data.get("percentComplete", 0)

            if status == "Succeeded":
                log.info("  Export complete (100%).")
                return
            if status == "Failed":
                err = data.get("error", {})
                raise RuntimeError(
                    f"Power BI export failed — "
                    f"code: {err.get('code', '?')} | "
                    f"message: {err.get('message', str(data))}"
                )
            log.info(f"  ... {status} ({pct}%)")
            time.sleep(POLL_INTERVAL)

        raise TimeoutError(
            f"Export did not finish within {EXPORT_TIMEOUT} seconds."
        )

    def _download(self, export_id: str, fpath: str) -> None:
        """Stream-download the completed export file to fpath."""
        url  = f"{PBI_API}/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}/exports/{export_id}/file"
        resp = requests.get(
            url, headers=self._api_headers(), timeout=120, stream=True
        )
        resp.raise_for_status()
        with open(fpath, "wb") as fh:
            for chunk in resp.iter_content(chunk_size=8_192):
                fh.write(chunk)


# ---------------------------------------------------------------------------
# Alias so main.py import ( from powerbi import PowerBIExporter ) keeps working
PowerBIExporter = PowerBIClient
