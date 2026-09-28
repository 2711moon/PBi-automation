"""
main.py -- Orchestrates the Power BI Report Automation.

Usage:
    python main.py          # Waits until SEND_AT time from config.py, then runs
    python main.py --now    # Runs immediately (for testing)

Flow:
    1. Wait until SEND_AT (skipped with --now)
    2. Prompt for Power BI email + password in terminal
    3. For each target group (AOM, Cluster Manager, ...):
         Phase 1: Export + email for every target in the group
    4. If any group had Phase-1 failures: ONE combined permission prompt,
       then Phase 2 (retry) -> Phase 3 (auto retry) per failing group
    5. Print per-phase + final consolidated summary per group
"""
import os
import sys
import time
import logging
import getpass
from collections import defaultdict
from datetime import datetime, date, timedelta

import openpyxl

from config import (
    STOREMASTER, EXPORTS_DIR, LOG_FILE,
    EMAIL_SUBJECT, EMAIL_BODY, SEND_AT, TARGET_GROUPS
)
from powerbi import PowerBIExporter
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


def resolve_date(keyword_or_literal):
    """
    Resolve a report_cfg date_from/date_to value to the "M/D/YYYY" string
    (no leading zeros) that the Date slicer's input fields expect.
      "yesterday"    -> today - 1 day
      "today"        -> today
      "month_start"  -> the 1st of the current month
      "YYYY-MM-DD"   -> that literal date
      None           -> None (date slicer left untouched)
    """
    if not keyword_or_literal:
        return None
    kw = keyword_or_literal.strip().lower()
    today_d = date.today()
    if kw == "yesterday":
        d = today_d - timedelta(days=1)
    elif kw == "today":
        d = today_d
    elif kw == "month_start":
        d = today_d.replace(day=1)
    else:
        d = datetime.strptime(keyword_or_literal.strip(), "%Y-%m-%d").date()
    return f"{d.month}/{d.day}/{d.year}"


def load_targets(group: dict) -> list:
    """
    Read Store Master.xlsx and return a list of unique targets for this
    group: {"name": ..., "delivery_email": ..., "columns": {col: value}}.

    `columns` holds every Store Master column referenced by the group's own
    value_column plus any report's filter_column (which can differ from
    value_column, e.g. "AOM Mail Id" for a report whose slicer searches by
    email instead of name) -- so export_for_target can look up whichever
    value a given report needs.
    """
    wb = openpyxl.load_workbook(STOREMASTER, data_only=True)
    ws = wb.active

    # Unhide all rows and columns unconditionally
    changed = False
    for rd in ws.row_dimensions.values():
        if rd.hidden:
            rd.hidden = False
            changed = True
    for cd in ws.column_dimensions.values():
        if cd.hidden:
            cd.hidden = False
            changed = True
    if changed:
        wb.save(STOREMASTER)
        wb = openpyxl.load_workbook(STOREMASTER, data_only=True)
        ws = wb.active

    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    value_column    = group["value_column"]
    delivery_column = group["delivery_column"]
    referenced_columns = {value_column}
    for report_cfg in group["reports"]:
        referenced_columns.add(report_cfg.get("filter_column", value_column))

    idx_value    = hdr.index(value_column)
    idx_delivery = hdr.index(delivery_column)
    idx_by_col   = {col: hdr.index(col) for col in referenced_columns}

    seen, targets = set(), []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name  = row[idx_value]
        dmail = row[idx_delivery]
        if not (name and dmail):
            continue
        name = str(name).strip()
        if name in seen:
            continue
        seen.add(name)

        columns = {}
        for col, idx in idx_by_col.items():
            v = row[idx]
            columns[col] = str(v).strip() if v else None

        targets.append({
            "name":           name,
            "delivery_email": str(dmail).strip(),
            "columns":        columns,
        })
    return targets


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


def print_phase_summary(phase_num: int, succeeded: list, failed: list):
    """Print a formatted summary block for a single phase."""
    w = 52
    log.info("")
    log.info("╔" + "═" * w + "╗")
    log.info(f"║  PHASE {phase_num} SUMMARY" + " " * (w - 16) + "║")
    log.info("╠" + "═" * w + "╣")

    if succeeded:
        names = ", ".join(succeeded)
        log.info(f"║  ✅ Sent     ({len(succeeded):>2}): {names[:w-18]}" +
                 (" ..." if len(names) > w - 18 else "") +
                 " " * max(0, w - 18 - min(len(names), w - 18)) + "  ║")
        # overflow lines
        remaining = names[w - 18:]
        while remaining:
            chunk = remaining[:w - 6]
            log.info(f"║      {chunk:<{w-6}}  ║")
            remaining = remaining[w - 6:]
    else:
        log.info(f"║  ✅ Sent     ( 0): —" + " " * (w - 22) + "  ║")

    if failed:
        names = ", ".join(failed)
        log.info(f"║  ❌ Failed   ({len(failed):>2}): {names[:w-18]}" +
                 (" ..." if len(names) > w - 18 else "") +
                 " " * max(0, w - 18 - min(len(names), w - 18)) + "  ║")
        remaining = names[w - 18:]
        while remaining:
            chunk = remaining[:w - 6]
            log.info(f"║      {chunk:<{w-6}}  ║")
            remaining = remaining[w - 6:]
    else:
        log.info(f"║  ❌ Failed   ( 0): —" + " " * (w - 22) + "  ║")

    log.info("╚" + "═" * w + "╝")
    log.info("")


def print_final_summary(group_name: str, total: int,
                        p1_ok: list, p1_fail: list,
                        p2_ok: list, p2_fail: list,
                        p3_ok: list, p3_fail: list):
    """Print the consolidated final summary across all phases for one group."""
    all_sent   = p1_ok + p2_ok + p3_ok
    all_failed = p3_fail if p3_fail is not None else (p2_fail if p2_fail is not None else p1_fail)
    w = 52

    log.info("")
    log.info("╔" + "═" * w + "╗")
    log.info(f"║  FINAL SUMMARY — {group_name}" + " " * max(0, w - 17 - len(group_name)) + "║")
    log.info("╠" + "═" * w + "╣")
    log.info(f"║  Total targets  : {total:<33}║")
    log.info(f"║  ✅ Emails sent : {len(all_sent):<33}║")
    log.info(f"║  ❌ Manual reqd : {len(all_failed):<33}║")

    if p2_ok or p2_fail is not None:
        log.info("╠" + "═" * w + "╣")
        log.info(f"║  Phase 1 sent: {len(p1_ok):<2}  Phase 1 failed: {len(p1_fail):<{w-32}}║")
        if p2_ok or p2_fail is not None:
            log.info(f"║  Phase 2 sent: {len(p2_ok):<2}  Phase 2 failed: {len(p2_fail):<{w-32}}║")
        if p3_ok or p3_fail is not None:
            log.info(f"║  Phase 3 sent: {len(p3_ok):<2}  Phase 3 failed: {len(p3_fail):<{w-32}}║")

    if all_failed:
        log.info("╠" + "═" * w + "╣")
        log.info(f"║  Please send these manually:" + " " * (w - 28) + "║")
        for name in all_failed:
            log.info(f"║    → {name:<{w-6}}║")
    log.info("╚" + "═" * w + "╝")
    log.info("")


# == Core export loop ==========================================================

def run_group_phase(exporter: "PowerBIExporter", mailer: "Mailer", group: dict,
                     targets: list, all_group_targets: list, today: str, phase_num: int) -> tuple:
    """
    Run a single export+email phase for the given targets within one group.

    Report-outer, target-inner: each report in the group is prepared
    (navigated to, date range + page set) exactly ONCE, then every target
    is looped applying just the slicer filter -- reusing the existing
    3-attempt search+scroll engine unchanged.

    A target is only counted as fully "succeeded" (for retry purposes) if
    it got a PDF from EVERY report in the group. A target that got at least
    one PDF but not all of them still receives an email (partial data beats
    no data) but is ALSO returned in `failed` so Phase 2/3 retries it.

    Returns (succeeded_names, failed_names) -- both plain lists of target names.
    """
    slicer_label = group["slicer_label"]
    reports      = group["reports"]

    pdfs_by_target        = defaultdict(list)
    report_success_count  = defaultdict(int)

    for report_cfg in reports:
        report_name = report_cfg["name"]

        resolved_cfg = dict(report_cfg)
        resolved_cfg["date_from"] = resolve_date(report_cfg.get("date_from"))
        resolved_cfg["date_to"]   = resolve_date(report_cfg.get("date_to"))

        log.info(f"\nPhase {phase_num} [{group['group_name']}] — Preparing report: {report_name}")
        try:
            exporter.prepare_report(resolved_cfg, slicer_label)
        except Exception as e:
            log.error(f"  Failed to prepare report '{report_name}': {e} -- skipping this report for all targets.")
            continue

        for target in targets:
            other_targets = [t for t in all_group_targets if t["name"] != target["name"]]
            log.info(f"  -> Exporting for: {target['name']}")

            pdf = exporter.export_for_target(target, resolved_cfg, slicer_label, other_targets, today)
            if pdf:
                pdfs_by_target[target["name"]].append(pdf)
                report_success_count[target["name"]] += 1
            else:
                log.error(f"  Export failed for report '{report_name}' -- skipping this attachment for {target['name']}")

    succeeded, failed = [], []
    for target in targets:
        name  = target["name"]
        dmail = target["delivery_email"]
        pdfs  = pdfs_by_target.get(name, [])

        if not pdfs:
            log.error(f"  All exports failed -- skipping email for {name}")
            failed.append(name)
            continue

        try:
            subject = EMAIL_SUBJECT.format(aom_name=name, date=today)
            body    = EMAIL_BODY.format(aom_name=name, date=today)
            mailer.send(dmail, subject, body, attachments=pdfs)
            log.info(f"  Email sent to {dmail} with {len(pdfs)} attachment(s)")
        except Exception as e:
            log.error(f"  Email failed for {name}: {e}")
            failed.append(name)
            continue

        if report_success_count[name] == len(reports):
            succeeded.append(name)
        else:
            failed.append(name)   # partial -- retry to try to complete it
            log.warning(
                f"  {name}: partial success, {report_success_count[name]}/{len(reports)} "
                f"reports -- emailed anyway, also queued for retry."
            )

    return succeeded, failed


# == Main ======================================================================

def main():
    run_now = "--now" in sys.argv

    os.makedirs(EXPORTS_DIR, exist_ok=True)
    cleanup_pdfs()

    log.info("=" * 55)
    log.info("  Power BI Report Automation")
    log.info(f"  {'Running immediately (--now)' if run_now else f'Scheduled at {SEND_AT}'}")
    log.info("=" * 55)

    if not run_now:
        wait_until_send_time()

    pbi_email, pbi_password = prompt_credentials()

    # Use yesterday's date for filenames/subject lines
    target_date = (date.today() - timedelta(days=1)).strftime("%d-%b-%Y")

    mailer = Mailer()

    # Load targets for every group up front
    group_targets = {}
    for group in TARGET_GROUPS:
        gname = group["group_name"]
        targets = load_targets(group)
        group_targets[gname] = targets
        log.info(f"Loaded {len(targets)} target(s) for group '{gname}' from Store Master.xlsx")
        for t in targets:
            log.info(f"  - {t['name']} | send to: {t['delivery_email']}")

    # Phase-result holders per group
    phase_results = {
        g["group_name"]: {"p1_ok": [], "p1_fail": [], "p2_ok": [], "p2_fail": None, "p3_ok": [], "p3_fail": None}
        for g in TARGET_GROUPS
    }

    with PowerBIExporter(pbi_email, pbi_password) as exporter:

        # ── Phase 1: every group, back-to-back ──────────────────────────────
        for group in TARGET_GROUPS:
            gname   = group["group_name"]
            targets = group_targets[gname]
            if not targets:
                log.info(f"\n  Group '{gname}': 0 targets loaded -- skipping.")
                continue

            log.info("\n" + "=" * 55)
            log.info(f"  PHASE 1 [{gname}] — Processing all targets")
            log.info("=" * 55)
            p1_ok, p1_fail = run_group_phase(exporter, mailer, group, targets, targets, target_date, phase_num=1)
            print_phase_summary(1, p1_ok, p1_fail)
            phase_results[gname]["p1_ok"]   = p1_ok
            phase_results[gname]["p1_fail"] = p1_fail

        # ── Phase 2: retry Phase 1 failures across all groups (ask ONE permission) ──
        failing_groups   = [g for g in TARGET_GROUPS if phase_results[g["group_name"]]["p1_fail"]]
        total_p1_failed  = sum(len(phase_results[g["group_name"]]["p1_fail"]) for g in failing_groups)

        if total_p1_failed:
            print(f"\n  {total_p1_failed} target(s) across {len(failing_groups)} group(s) failed in Phase 1.")
            answer = input("  Run Phase 2 to retry them? [Y/n]: ").strip().lower()
            if answer in ("", "y", "yes"):
                for group in failing_groups:
                    gname   = group["group_name"]
                    targets = group_targets[gname]
                    p1_fail = phase_results[gname]["p1_fail"]
                    retry_targets = [t for t in targets if t["name"] in p1_fail]

                    log.info("\n" + "=" * 55)
                    log.info(f"  PHASE 2 [{gname}] — Retrying failed targets")
                    log.info("=" * 55)
                    log.info("  Re-logging in for a fresh session before Phase 2...")
                    exporter.relogin()
                    p2_ok, p2_fail = run_group_phase(exporter, mailer, group, retry_targets, targets, target_date, phase_num=2)
                    print_phase_summary(2, p2_ok, p2_fail)
                    phase_results[gname]["p2_ok"]   = p2_ok
                    phase_results[gname]["p2_fail"] = p2_fail

                    # ── Phase 3: retry Phase 2 failures (automatic) ─────────
                    if p2_fail:
                        retry_targets = [t for t in targets if t["name"] in p2_fail]
                        log.info("\n" + "=" * 55)
                        log.info(f"  PHASE 3 [{gname}] — Final retry (automatic)")
                        log.info("=" * 55)
                        log.info("  Re-logging in for a fresh session before Phase 3...")
                        exporter.relogin()
                        p3_ok, p3_fail = run_group_phase(exporter, mailer, group, retry_targets, targets, target_date, phase_num=3)
                        print_phase_summary(3, p3_ok, p3_fail)
                        phase_results[gname]["p3_ok"]   = p3_ok
                        phase_results[gname]["p3_fail"] = p3_fail
                    else:
                        log.info(f"  Group '{gname}': Phase 2 achieved 100% success — Phase 3 not needed.")
            else:
                log.info("  Phase 2 skipped by user.")
        else:
            log.info("  Phase 1 achieved 100% success across all groups — Phase 2 and Phase 3 not needed.")

    # ── Cleanup + final summary (one block per group) ────────────────────────
    cleanup_pdfs()
    for group in TARGET_GROUPS:
        gname = group["group_name"]
        targets = group_targets[gname]
        if not targets:
            continue
        r = phase_results[gname]
        print_final_summary(gname, len(targets), r["p1_ok"], r["p1_fail"],
                             r["p2_ok"], r["p2_fail"], r["p3_ok"], r["p3_fail"])


if __name__ == "__main__":
    main()
