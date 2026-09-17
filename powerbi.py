"""
powerbi.py  —  Option B: browser-native API calls via Playwright page.evaluate().

AUTHENTICATION : Playwright headless browser (email + password, already working).
EXPORT         : ALL Power BI API calls (ExportTo, poll, download) are made
                 via JavaScript fetch() INSIDE the browser session.
                 No token extracted to Python. No Azure AD app needed.
                 effectiveIdentity applied server-side → correct RLS per AOM.

Diagnostic:
  If the very first API call (get dataset ID) returns 401 even from the browser,
  the log will say "TOKEN SCOPE ERROR" and tell you Azure AD app registration
  is the only remaining fix. That result in 30 seconds — no wasted time.
"""
import os
import base64
import time
import logging
from typing import Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from config import (
    PBI_GROUP_ID, PBI_REPORT_ID, PBI_RLS_ROLE,
    EXPORTS_DIR,
)

log = logging.getLogger(__name__)

NAV_TIMEOUT    = 60_000   # ms  — browser navigation
EVAL_TIMEOUT   = 120_000  # ms  — JS evaluate (long for PDF download)
POLL_INTERVAL  = 6        # sec — between export-status polls
EXPORT_TIMEOUT = 300      # sec — give up if export takes longer
PBI_API        = "https://api.powerbi.com/v1.0/myorg"

# JavaScript that finds the Power BI Bearer token in localStorage/sessionStorage
_JS_GET_TOKEN = """
    (() => {
        for (const store of [localStorage, sessionStorage]) {
            for (let i = 0; i < store.length; i++) {
                const key = store.key(i) || '';
                try {
                    const item = JSON.parse(store.getItem(key));
                    if (item && item.secret) {
                        const t = (item.target || '').toLowerCase();
                        if (t.includes('powerbi') || t.includes('analysis.windows.net')) {
                            return item.secret;
                        }
                    }
                } catch(e) {}
            }
        }
        return null;
    })()
"""


class PowerBIClient:
    """
    Context manager.  Browser stays open for the entire run so all
    API calls share the same authenticated session.
    """

    def __init__(self, email: str, password: str):
        self.email       = email
        self.password    = password
        self._pw         = None
        self._browser    = None
        self._page       = None
        self._dataset_id: Optional[str] = None

    # ------------------------------------------------------------------ lifecycle

    def __enter__(self) -> "PowerBIClient":
        log.info("Launching headless browser...")
        self._pw      = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )
        ctx = self._browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/127.0.0.0 Safari/537.36"
            ),
        )
        self._page = ctx.new_page()
        self._page.set_default_timeout(EVAL_TIMEOUT)

        self._do_login(self._page)

        log.info("  Navigating to Power BI home (loading session tokens)...")
        self._page.goto(
            "https://app.powerbi.com/",
            timeout=NAV_TIMEOUT,
            wait_until="domcontentloaded",
        )
        self._page.wait_for_timeout(6_000)   # let MSAL.js cache tokens

        log.info("  Fetching dataset ID via browser fetch()...")
        self._dataset_id = self._get_dataset_id()
        log.info(f"  Dataset ID: {self._dataset_id}")
        return self

    def __exit__(self, *args) -> None:
        try:
            if self._browser:
                self._browser.close()
            if self._pw:
                self._pw.stop()
        except Exception:
            pass
        log.info("Browser closed.")

    # ------------------------------------------------------------------ login

    def _do_login(self, page) -> None:
        log.info("  Navigating to Microsoft login...")
        page.goto(
            "https://login.microsoftonline.com/",
            timeout=NAV_TIMEOUT,
            wait_until="domcontentloaded",
        )
        page.wait_for_timeout(1_500)

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

        for sel in ['input[name="passwd"]', 'input[type="password"]', "#i0118"]:
            try:
                page.wait_for_selector(sel, timeout=10_000)
                page.fill(sel, self.password)
                page.keyboard.press("Enter")
                log.info("  Password submitted.")
                break
            except PWTimeout:
                continue
        page.wait_for_timeout(2_000)

        self._handle_mfa(page)

        try:
            page.wait_for_selector("#idBtn_Back", timeout=8_000)
            page.click("#idBtn_Back")
            log.info("  'Stay signed in?' -- No.")
        except PWTimeout:
            pass

        log.info("  Login complete.")

    def _handle_mfa(self, page) -> None:
        try:
            page.wait_for_selector(
                'input[name="otc"], #idTxtBx_SAOTCC_OTC', timeout=8_000
            )
            otp = input("\n  MFA code required: ").strip()
            page.fill('input[name="otc"], #idTxtBx_SAOTCC_OTC', otp)
            page.keyboard.press("Enter")
            page.wait_for_timeout(3_000)
            return
        except PWTimeout:
            pass

        try:
            page.wait_for_selector(
                '[data-value="PhoneAppNotification"], #idDiv_SAOTCS_Section',
                timeout=5_000,
            )
            log.info("  Authenticator push sent — approve on phone (20s)...")
            page.wait_for_timeout(20_000)
            return
        except PWTimeout:
            pass

        log.info("  No MFA challenge detected.")

    # ------------------------------------------------------------------ browser API helpers

    def _browser_fetch(self, method: str, url: str, body: dict = None) -> dict:
        """
        Make an authenticated REST API call from inside the browser JS context.
        Returns { status: <int>, data: <dict> } or { __error: <str> }.
        """
        result = self._page.evaluate(
            """async (p) => {
                // Find Power BI Bearer token in browser storage
                let token = null;
                for (const store of [localStorage, sessionStorage]) {
                    for (let i = 0; i < store.length; i++) {
                        const key = store.key(i) || '';
                        try {
                            const item = JSON.parse(store.getItem(key));
                            if (item && item.secret) {
                                const t = (item.target || '').toLowerCase();
                                if (t.includes('powerbi') ||
                                    t.includes('analysis.windows.net')) {
                                    token = item.secret;
                                    break;
                                }
                            }
                        } catch(e) {}
                    }
                    if (token) break;
                }

                if (!token) return { __error: 'NO_TOKEN',
                    detail: 'No Power BI token found in localStorage/sessionStorage' };

                const opts = {
                    method: p.method,
                    headers: {
                        'Authorization': 'Bearer ' + token,
                        'Content-Type': 'application/json',
                    },
                };
                if (p.body) opts.body = JSON.stringify(p.body);

                try {
                    const resp = await fetch(p.url, opts);
                    const ct   = resp.headers.get('content-type') || '';
                    const data = ct.includes('json')
                        ? await resp.json()
                        : { __text: await resp.text() };
                    return { status: resp.status, data };
                } catch(err) {
                    return { __error: 'FETCH_ERROR', detail: String(err) };
                }
            }""",
            {"method": method, "url": url, "body": body},
        )
        return result

    def _browser_download(self, url: str) -> bytes:
        """Download binary file from browser, return as Python bytes via base64."""
        result = self._page.evaluate(
            """async (p) => {
                let token = null;
                for (const store of [localStorage, sessionStorage]) {
                    for (let i = 0; i < store.length; i++) {
                        const key = store.key(i) || '';
                        try {
                            const item = JSON.parse(store.getItem(key));
                            if (item && item.secret) {
                                const t = (item.target || '').toLowerCase();
                                if (t.includes('powerbi') ||
                                    t.includes('analysis.windows.net')) {
                                    token = item.secret;
                                    break;
                                }
                            }
                        } catch(e) {}
                    }
                    if (token) break;
                }
                if (!token) return { __error: 'NO_TOKEN' };

                try {
                    const resp = await fetch(p.url, {
                        headers: { 'Authorization': 'Bearer ' + token }
                    });
                    if (!resp.ok) {
                        return { __error: 'HTTP_' + resp.status,
                                 detail: await resp.text() };
                    }
                    const blob = await resp.blob();
                    return new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload  = () => resolve(
                            { base64: reader.result.split(',')[1] }
                        );
                        reader.onerror = () => reject('FileReader error');
                        reader.readAsDataURL(blob);
                    });
                } catch(err) {
                    return { __error: 'FETCH_ERROR', detail: String(err) };
                }
            }""",
            {"url": url},
        )
        if "__error" in result:
            raise RuntimeError(f"Download failed: {result}")
        return base64.b64decode(result["base64"])

    # ------------------------------------------------------------------ dataset ID

    def _get_dataset_id(self) -> str:
        url    = f"{PBI_API}/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}"
        result = self._browser_fetch("GET", url)

        if "__error" in result:
            raise RuntimeError(
                f"Could not call Power BI API from browser: {result}\n"
                "Check that the browser is on app.powerbi.com and logged in."
            )

        if result["status"] == 401:
            raise RuntimeError(
                "TOKEN SCOPE ERROR — API returned 401 even from inside the browser.\n"
                "This means the localStorage token's audience does not match the\n"
                "Power BI REST API. Azure AD app registration is required.\n"
                "Admin guide: https://aad.portal.azure.com → App registrations\n"
                f"Raw response: {result.get('data', '')}"
            )

        if result["status"] != 200:
            raise RuntimeError(
                f"GetReport failed: HTTP {result['status']} — {result.get('data', '')}"
            )

        return result["data"]["datasetId"]

    # ------------------------------------------------------------------ per-AOM export

    def export_report(
        self, filter_email: str, aom_name: str, date_str: str
    ) -> Optional[str]:
        """
        Export report for one AOM via browser fetch() with effectiveIdentity.
        Returns local PDF path or None on failure.
        """
        os.makedirs(EXPORTS_DIR, exist_ok=True)
        safe  = "".join(c if c.isalnum() or c in " _-" else "_" for c in aom_name)
        fname = f"{safe.replace(' ', '_')}_{date_str}.pdf"
        fpath = os.path.join(EXPORTS_DIR, fname)

        try:
            log.info(f"  Starting export (effectiveIdentity: {filter_email})...")
            export_id = self._start_export(filter_email)

            log.info("  Waiting for Power BI to render PDF...")
            self._poll_export(export_id)

            log.info("  Downloading PDF via browser...")
            pdf_bytes = self._browser_download(
                f"{PBI_API}/groups/{PBI_GROUP_ID}"
                f"/reports/{PBI_REPORT_ID}/exports/{export_id}/file"
            )

            with open(fpath, "wb") as fh:
                fh.write(pdf_bytes)

            log.info(f"  Saved: {fname}  ({len(pdf_bytes):,} bytes)")
            return fpath

        except Exception as exc:
            log.error(f"  Export failed for {aom_name}: {exc}")
            return None

    def _start_export(self, filter_email: str) -> str:
        url    = f"{PBI_API}/groups/{PBI_GROUP_ID}/reports/{PBI_REPORT_ID}/ExportTo"
        body   = {
            "format": "PDF",
            "powerBIReportConfiguration": {
                "identities": [{
                    "username": filter_email,
                    "roles":    [PBI_RLS_ROLE],
                    "datasets": [self._dataset_id],
                }]
            },
        }
        result = self._browser_fetch("POST", url, body)

        if result.get("status") not in (200, 202):
            raise RuntimeError(
                f"ExportTo failed: HTTP {result.get('status')} — {result.get('data', '')}"
            )
        return result["data"]["id"]

    def _poll_export(self, export_id: str) -> None:
        url      = (
            f"{PBI_API}/groups/{PBI_GROUP_ID}"
            f"/reports/{PBI_REPORT_ID}/exports/{export_id}"
        )
        deadline = time.time() + EXPORT_TIMEOUT

        while time.time() < deadline:
            result = self._browser_fetch("GET", url)

            if result.get("status") != 200:
                raise RuntimeError(
                    f"Poll failed: HTTP {result.get('status')} — {result.get('data', '')}"
                )

            data   = result["data"]
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

        raise TimeoutError(f"Export timed out after {EXPORT_TIMEOUT}s.")


# Alias — keeps main.py import unchanged
PowerBIExporter = PowerBIClient
