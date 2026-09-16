"""
main.py -- Orchestrates the Power BI Report Automation.

Usage:
    python main.py          # Waits until SEND_AT time from config.py, then runs
    python main.py --now    # Runs immediately (for testing)

Flow:
    1. Wait until SEND_AT (skipped with --now)
    2. Prompt for Power BI admin email + password in terminal
    3. Log into Power BI via headless browser — handles MFA in terminal
    4. Capture Bearer token from intercepted API requests, then close browser
    5. For each AOM in StoreMAster.xlsx:
       a. Call Power BI REST Export API with effectiveIdentity (RLS applied server-side)
       b. Poll until PDF is ready, download it
       c. Email the PDF via Zoho SMTP
       d. Delete the local PDF
"""
import os
import sys
import time
import logging
import getpass
from datetime import datetime, date

import openpyxl

from config import (
    STOREMASTER, EXPORTS_DIR, LOG_FILE,
    EMAIL_SUBJECT, EMAIL_BODY, SEND_AT,
    PBI_WORKSPACE_NAME, PBI_REPORT_NAME,
)
from powerbi import PowerBIExporter   # PowerBIExporter = PowerBIClient alias
from mailer import Mailer

# == Logging ===================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# == Helpers ===================================================================

def wait_until_send_time():
    """Block until SEND_AT time is reached. Shows a simple countdown."""
    now    = datetime.now()
    hh, mm = (int(x) for x in SEND_AT.split(":"))
    target = now.replace(hour=hh, minute=mm, second=0, microsecond=0)

    if target <= now:
        log.info(f"Send time {SEND_AT} has already passed today -- running immediately.")
        return

    wait_secs = (target - now).total_seconds()
    log.info(f"Scheduled to run at {SEND_AT}. Waiting {int(wait_secs // 60)} min {int(wait_secs % 60)} sec...")
    log.info("(Run  python main.py --now  to skip the wait and test immediately.)")

    while True:
        remaining = (target - datetime.now()).total_seconds()
        if remaining <= 0:
            break
        # Print a dot every 60 seconds so the terminal shows the script is alive
        if int(remaining) % 60 == 0:
            print(f"  ... {int(remaining // 60)} min remaining until {SEND_AT}", flush=True)
        time.sleep(1)

    log.info(f"Send time reached. Starting automation.")


def prompt_credentials():
    """Ask for Power BI email and password in the terminal."""
    print()
    print("=" * 55)
    print("  Power BI Login")
    print("=" * 55)
    email    = input("  Email    : ").strip()
    password = getpass.getpass("  Password : ")
    print("=" * 55)
    print()
    if not email or not password:
        log.error("Email or password cannot be empty.")
        sys.exit(1)
    return email, password


def load_aoms():
    """Read StoreMAster.xlsx and return a list of unique AOMs."""
    wb  = openpyxl.load_workbook(STOREMASTER)
    ws  = wb.active
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    idx_aom   = hdr.index("AOM")
    idx_fmail = hdr.index("AOM Mail Id")   # RLS filter value
    idx_email = hdr.index("Email")          # delivery address

    seen, aoms = set(), []
    for row in ws.iter_rows(min_row=2, values_only=True):
        aom_name  = row[idx_aom]
        aom_mail  = row[idx_fmail]
        del_email = row[idx_email]
        if not (aom_name and aom_mail and del_email):
            continue
        if aom_mail in seen:
            continue
        seen.add(aom_mail)
        aoms.append({
            "name":           str(aom_name).strip(),
            "filter_email":   str(aom_mail).strip(),
            "delivery_email": str(del_email).strip(),
        })
    return aoms


def cleanup_pdfs():
    """Delete all PDFs from exports/ folder."""
    if not os.path.exists(EXPORTS_DIR):
        return
    for f in os.listdir(EXPORTS_DIR):
        if f.lower().endswith(".pdf"):
            try:
                os.remove(os.path.join(EXPORTS_DIR, f))
            except OSError:
                pass


# == Main ======================================================================

def main():
    run_now = "--now" in sys.argv

    os.makedirs(EXPORTS_DIR, exist_ok=True)
    cleanup_pdfs()

    log.info("=" * 55)
    log.info("  Power BI Report Automation")
    log.info(f"  {'Running immediately (--now)' if run_now else f'Scheduled at {SEND_AT}'}")
    log.info("=" * 55)

    # Step 1: Wait until send time (unless --now)
    if not run_now:
        wait_until_send_time()

    # Step 2: Ask for credentials at send time
    pbi_email, pbi_password = prompt_credentials()

    # Step 3: Load AOMs
    aoms = load_aoms()
    today = date.today().strftime("%d-%b-%Y")
    log.info(f"Loaded {len(aoms)} AOM(s) from StoreMAster.xlsx")
    for a in aoms:
        log.info(f"  - {a['name']} | filter: {a['filter_email']} | send to: {a['delivery_email']}")

    # Step 4 & 5: Authenticate, export via REST API with effectiveIdentity, email
    mailer        = Mailer()
    success, fail = 0, 0

    with PowerBIExporter(pbi_email, pbi_password) as exporter:
        for aom in aoms:
            name  = aom["name"]
            fmail = aom["filter_email"]
            dmail = aom["delivery_email"]

            log.info(f"\nProcessing: {name}")

            pdf = exporter.export_report(
                filter_email=fmail,
                aom_name=name,
                date_str=today,
            )

            if not pdf:
                log.error(f"  Export failed -- skipping email for {name}")
                fail += 1
                continue

            try:
                subject = EMAIL_SUBJECT.format(aom_name=name, date=today)
                body    = EMAIL_BODY.format(aom_name=name, date=today)
                mailer.send(dmail, subject, body, pdf)
                log.info(f"  Email sent to {dmail}")
                success += 1
            except Exception as e:
                log.error(f"  Email failed for {name}: {e}")
                fail += 1

    # Step 6: Clean up
    cleanup_pdfs()

    log.info("\n" + "=" * 55)
    log.info(f"  Done.  Sent: {success}   Failed: {fail}")
    log.info("=" * 55)


if __name__ == "__main__":
    main()
