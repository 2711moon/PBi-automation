"""
main.py -- Orchestrates the Power BI Report Automation.

Usage:
    python main.py          # Waits until SEND_AT time from config.py, then runs
    python main.py --now    # Runs immediately (for testing)

Flow:
    1. Wait until SEND_AT (skipped with --now)
    2. Prompt for Power BI email + password in terminal
    3. For each target group (AOM, Cluster Manager, ...), Phase 1 gathers
       PDFs for every target. Targets that got every report are emailed
       immediately; targets missing at least one report are held (NOT
       emailed yet).
    4. If any targets are held (across any group), ONE combined prompt asks:
         1. Send partial mail now (whatever data is available)
         2. Run Phase 2 to try to complete them first
       Choosing 2 retries held targets (Phase 2, then Phase 3 automatically
       if still incomplete) and only then sends -- so nobody gets more than
       one email.
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


def print_phase_summary(phase_num: int, complete: list, incomplete: list):
    """
    Print a formatted summary block for a single gather phase.
    NOTE: "complete"/"incomplete" describe whether a target got every
    report in this phase -- NOT whether an email was sent (sending is
    decoupled from gathering; see gather_phase()/send_for_targets()).
    """
    w = 52
    log.info("")
    log.info("╔" + "═" * w + "╗")
    log.info(f"║  PHASE {phase_num} SUMMARY" + " " * (w - 16) + "║")
    log.info("╠" + "═" * w + "╣")

    if complete:
        names = ", ".join(complete)
        log.info(f"║  ✅ Complete ({len(complete):>2}): {names[:w-18]}" +
                 (" ..." if len(names) > w - 18 else "") +
                 " " * max(0, w - 18 - min(len(names), w - 18)) + "  ║")
        # overflow lines
        remaining = names[w - 18:]
        while remaining:
            chunk = remaining[:w - 6]
            log.info(f"║      {chunk:<{w-6}}  ║")
            remaining = remaining[w - 6:]
    else:
        log.info(f"║  ✅ Complete ( 0): —" + " " * (w - 22) + "  ║")

    if incomplete:
        names = ", ".join(incomplete)
        log.info(f"║  ⏳ Incomplete({len(incomplete):>2}): {names[:w-18]}" +
                 (" ..." if len(names) > w - 18 else "") +
                 " " * max(0, w - 18 - min(len(names), w - 18)) + "  ║")
        remaining = names[w - 18:]
        while remaining:
            chunk = remaining[:w - 6]
            log.info(f"║      {chunk:<{w-6}}  ║")
            remaining = remaining[w - 6:]
    else:
        log.info(f"║  ⏳ Incomplete( 0): —" + " " * (w - 22) + "  ║")

    log.info("╚" + "═" * w + "╝")
    log.info("")


def print_final_summary(group_name: str, total: int, sent: list, failed: list):
    """Print the consolidated final summary (actual email dispatch) for one group."""
    w = 52

    log.info("")
    log.info("╔" + "═" * w + "╗")
    log.info(f"║  FINAL SUMMARY — {group_name}" + " " * max(0, w - 17 - len(group_name)) + "║")
    log.info("╠" + "═" * w + "╣")
    log.info(f"║  Total targets  : {total:<33}║")
    log.info(f"║  ✅ Emails sent : {len(sent):<33}║")
    log.info(f"║  ❌ Manual reqd : {len(failed):<33}║")

    if failed:
        log.info("╠" + "═" * w + "╣")
        log.info(f"║  Please send these manually:" + " " * (w - 28) + "║")
        for name in failed:
            log.info(f"║    → {name:<{w-6}}║")
    log.info("╚" + "═" * w + "╝")
    log.info("")


# == Core export loop ==========================================================

def gather_phase(exporter: "PowerBIExporter", group: dict, targets: list,
                  all_group_targets: list, today: str, phase_num: int) -> dict:
    """
    Gather PDFs for the given targets within one group -- does NOT send any
    email (sending is a separate, deliberate step; see send_for_targets()).

    Report-outer, target-inner: each report in the group is prepared
    (navigated to, date range + page set) exactly ONCE, then every target
    is looped applying just the slicer filter -- reusing the existing
    3-attempt search+scroll engine unchanged.

    Returns {target_name: {"pdfs": [...], "complete": bool}}, where
    "complete" means the target got a PDF from EVERY report in the group.
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

    result = {}
    for target in targets:
        name = target["name"]
        result[name] = {
            "pdfs":     pdfs_by_target.get(name, []),
            "complete": report_success_count.get(name, 0) == len(reports),
        }
    return result


def send_for_targets(mailer: "Mailer", targets_by_name: dict, gather_result: dict,
                      today: str, sent_names: list, failed_names: list) -> None:
    """
    Send one email per target present in `gather_result`, using whatever
    PDFs it has (full or partial). Targets with zero PDFs get no email and
    are recorded as failed instead.

    `sent_names`/`failed_names` are mutated in place so callers can
    accumulate results across multiple calls (immediate sends for
    fully-complete targets at each phase, plus a final send for
    partial/terminal targets) into one running tally per group.
    """
    for name, info in gather_result.items():
        pdfs   = info["pdfs"]
        target = targets_by_name[name]

        if not pdfs:
            log.error(f"  All exports failed -- skipping email for {name}")
            failed_names.append(name)
            continue

        try:
            subject = EMAIL_SUBJECT.format(aom_name=name, date=today)
            body    = EMAIL_BODY.format(aom_name=name, date=today)
            mailer.send(target["delivery_email"], subject, body, attachments=pdfs)
            log.info(f"  Email sent to {target['delivery_email']} with {len(pdfs)} attachment(s)")
            sent_names.append(name)
        except Exception as e:
            log.error(f"  Email failed for {name}: {e}")
            failed_names.append(name)


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
    group_targets_by_name = {}
    for group in TARGET_GROUPS:
        gname = group["group_name"]
        targets = load_targets(group)
        group_targets[gname] = targets
        group_targets_by_name[gname] = {t["name"]: t for t in targets}
        log.info(f"Loaded {len(targets)} target(s) for group '{gname}' from Store Master.xlsx")
        for t in targets:
            log.info(f"  - {t['name']} | send to: {t['delivery_email']}")

    # Running per-group tallies of what actually got emailed vs. needs manual
    # follow-up -- built up across however many phases each group goes through.
    sent_by_group   = {g["group_name"]: [] for g in TARGET_GROUPS}
    failed_by_group = {g["group_name"]: [] for g in TARGET_GROUPS}
    # Targets still missing at least one report after Phase 1, held (NOT
    # emailed yet) until we know whether Phase 2 will run.
    pending_by_group     = {}
    phase1_result_by_group = {}

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
            result = gather_phase(exporter, group, targets, targets, target_date, phase_num=1)
            complete_names   = [n for n, i in result.items() if i["complete"]]
            incomplete_names = [n for n, i in result.items() if not i["complete"]]
            print_phase_summary(1, complete_names, incomplete_names)
            phase1_result_by_group[gname] = result

            # Complete targets have nothing left to decide -- send now.
            send_for_targets(
                mailer, group_targets_by_name[gname],
                {n: result[n] for n in complete_names},
                target_date, sent_by_group[gname], failed_by_group[gname]
            )
            pending_by_group[gname] = incomplete_names

        # ── Decide what to do with incomplete targets, across ALL groups ────
        total_pending  = sum(len(v) for v in pending_by_group.values())
        pending_groups = sum(1 for v in pending_by_group.values() if v)

        if total_pending:
            print(f"\n  {total_pending} target(s) across {pending_groups} group(s) are incomplete (missing at least one report).")
            choice = None
            while choice not in ("1", "2"):
                print("\n  What would you like to do?")
                print("    1. Send partial mail now (whatever data is available)")
                print("    2. Run Phase 2 to try to complete them first")
                choice = input("\n  Choice [1/2]: ").strip()

            if choice == "1":
                log.info("  Sending partial mail for incomplete targets (no retry requested).")
                for group in TARGET_GROUPS:
                    gname = group["group_name"]
                    pending_names = pending_by_group.get(gname) or []
                    if not pending_names:
                        continue
                    result = phase1_result_by_group[gname]
                    send_for_targets(
                        mailer, group_targets_by_name[gname],
                        {n: result[n] for n in pending_names},
                        target_date, sent_by_group[gname], failed_by_group[gname]
                    )
            else:
                # ── Phase 2: retry pending targets, per group ────────────────
                for group in TARGET_GROUPS:
                    gname = group["group_name"]
                    pending_names = pending_by_group.get(gname) or []
                    if not pending_names:
                        continue

                    targets_by_name = group_targets_by_name[gname]
                    all_targets     = group_targets[gname]
                    retry_targets   = [targets_by_name[n] for n in pending_names]

                    log.info("\n" + "=" * 55)
                    log.info(f"  PHASE 2 [{gname}] — Retrying incomplete targets")
                    log.info("=" * 55)
                    log.info("  Re-logging in for a fresh session before Phase 2...")
                    exporter.relogin()
                    result2 = gather_phase(exporter, group, retry_targets, all_targets, target_date, phase_num=2)
                    complete2   = [n for n, i in result2.items() if i["complete"]]
                    incomplete2 = [n for n, i in result2.items() if not i["complete"]]
                    print_phase_summary(2, complete2, incomplete2)

                    # Newly-complete targets have nothing left to decide -- send now.
                    send_for_targets(
                        mailer, targets_by_name, {n: result2[n] for n in complete2},
                        target_date, sent_by_group[gname], failed_by_group[gname]
                    )

                    if incomplete2:
                        # ── Phase 3: final automatic retry, then send everyone
                        # remaining regardless of completeness -- terminal, no
                        # further phase to hold out for.
                        log.info("\n" + "=" * 55)
                        log.info(f"  PHASE 3 [{gname}] — Final retry (automatic)")
                        log.info("=" * 55)
                        log.info("  Re-logging in for a fresh session before Phase 3...")
                        exporter.relogin()
                        retry_targets3 = [targets_by_name[n] for n in incomplete2]
                        result3 = gather_phase(exporter, group, retry_targets3, all_targets, target_date, phase_num=3)
                        complete3   = [n for n, i in result3.items() if i["complete"]]
                        incomplete3 = [n for n, i in result3.items() if not i["complete"]]
                        print_phase_summary(3, complete3, incomplete3)
                        send_for_targets(
                            mailer, targets_by_name, result3,
                            target_date, sent_by_group[gname], failed_by_group[gname]
                        )
                    else:
                        log.info(f"  Group '{gname}': Phase 2 achieved 100% success — Phase 3 not needed.")
        else:
            log.info("  Phase 1 achieved 100% success across all groups — nothing pending.")

    # ── Cleanup + final summary (one block per group) ────────────────────────
    cleanup_pdfs()
    for group in TARGET_GROUPS:
        gname = group["group_name"]
        targets = group_targets[gname]
        if not targets:
            continue
        print_final_summary(gname, len(targets), sent_by_group[gname], failed_by_group[gname])


if __name__ == "__main__":
    main()
