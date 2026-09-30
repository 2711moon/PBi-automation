"""
main.py -- Orchestrates the Power BI Report Automation.

Usage:
    python main.py          # Waits until SEND_AT time from config.py, then runs
    python main.py --now    # Runs immediately (for testing)

Flow:
    1. Wait until SEND_AT (skipped with --now)
    2. Prompt for Power BI email + password in terminal
    3. Phase 1 gathers one PDF per (target, report) pair, for every group.
       Every pair is individually classified as:
         - success   : found, no conflicting name -- always included.
         - Group A   : not found / unstable across the 2 attempts, OR
                       found but with 3+ conflicting names (too messy to
                       be a plausible coincidence) -- needs RETRY or DROP.
         - Group B   : found, with exactly 1-2 conflicting names (a
                       plausible same-name-different-designation
                       coincidence) -- needs GO AHEAD or DROP.
       A target whose every report is a clean "success" is emailed
       immediately after Phase 1 -- nothing to review.
    4. For everything else, Group A and Group B are shown per (target,
       report) pair, not per target -- so a person present in 4 of 5
       reports only gets asked about the 1 that actually needs a decision.
       RETRY re-runs just that (target, report) pair (Phase 2, fresh
       login), then the SAME classifier is applied to the new result:
       success -> included; Group B -> one more go-ahead/drop question;
       Group A again -> excluded, final, no further retry offered.
       DROP only excludes that one report's attachment -- every other
       report already resolved for that person is unaffected.
    5. Final send: every target with at least one included report gets
       one email with whatever PDFs are included (full or partial).
       Zero included reports -> no email, listed as needing manual send.
    6. Every phase prints a per-target/per-report status table, plus a
       final consolidated summary per group.
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

# 1-2 distinct colliding names on a "conflict" pair is treated as a
# plausible same-name-different-designation coincidence (Group B, a human
# eyeball call); 3+ is treated as too messy to be ordinary chance (Group A,
# retry-or-drop instead).
MAX_OTHER_CONFLICTS_FOR_GROUP_B = 2


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


def classify_pair(status: str, conflicts: list) -> str:
    """
    Classify one (target, report)'s verification outcome into
    "success" | "A" | "B":
      - success (found, 0 conflicts)                      -> "success"
      - not_found / dual (unstable across the 2 attempts)  -> "A"
      - conflict, 1-2 distinct other names                  -> "B"
      - conflict, 3+ distinct other names                   -> "A"
    """
    if status == "success":
        return "success"
    if status == "conflict":
        return "B" if len(conflicts) <= MAX_OTHER_CONFLICTS_FOR_GROUP_B else "A"
    return "A"  # not_found, dual


# == Core export loop ==========================================================

def gather_for_pairs(exporter: "PowerBIExporter", group: dict, pairs: list,
                      all_group_targets: list, today: str, label: str) -> dict:
    """
    Run export_for_target for exactly the given (target, report_cfg) pairs,
    grouped by report so each report is only prepare_report()'d once --
    used both for Phase 1 (every target x every report) and for a scoped
    retry (only the specific failed (target, report) pairs).

    Returns {(target_name, report_name): {"pdf":, "status":, "conflicts":}}.
    """
    slicer_label = group["slicer_label"]

    by_report = defaultdict(list)
    report_cfg_by_name = {}
    for target, report_cfg in pairs:
        by_report[report_cfg["name"]].append(target)
        report_cfg_by_name[report_cfg["name"]] = report_cfg

    results = {}
    for report_name, targets_for_report in by_report.items():
        report_cfg = report_cfg_by_name[report_name]
        resolved_cfg = dict(report_cfg)
        resolved_cfg["date_from"] = resolve_date(report_cfg.get("date_from"))
        resolved_cfg["date_to"]   = resolve_date(report_cfg.get("date_to"))

        log.info(f"\n{label} [{group['group_name']}] -- Preparing report: {report_name}")
        try:
            exporter.prepare_report(resolved_cfg, slicer_label)
        except Exception as e:
            log.error(f"  Failed to prepare report '{report_name}': {e} -- skipping this report for all targets.")
            for target in targets_for_report:
                results[(target["name"], report_name)] = {"pdf": None, "status": "not_found", "conflicts": []}
            continue

        for target in targets_for_report:
            other_targets = [t for t in all_group_targets if t["name"] != target["name"]]
            log.info(f"  -> Exporting for: {target['name']}")
            pdf, status, conflicts = exporter.export_for_target(
                target, resolved_cfg, slicer_label, other_targets, today
            )
            results[(target["name"], report_name)] = {"pdf": pdf, "status": status, "conflicts": conflicts}
            if pdf:
                log.info(f"  [{target['name']} / {report_name}] result: {status} (PDF saved)")
            else:
                log.error(f"  [{target['name']} / {report_name}] result: {status} (no PDF)")

    return results


def print_status_table(title: str, rows: list) -> None:
    """
    rows: list of (target_name, report_name, status_label) already-
    formatted strings -- a simple, correctness-first table (layout doesn't
    matter, per-row data does).
    """
    log.info("")
    log.info(f"-- {title} --")
    if not rows:
        log.info("  (none)")
    for target_name, report_name, status_label in rows:
        log.info(f"  {target_name:<30} | {report_name:<30} | {status_label}")
    log.info("")


def prompt_group_a(entries: list) -> dict:
    """
    entries: list of (target_name, report_name, reason) for Group A pairs.
    Returns {(target_name, report_name): "retry"|"drop"}.
    """
    decisions = {}
    if not entries:
        return decisions
    print("\n  Group A -- not found / unstable even after all attempts:")
    for target_name, report_name, reason in entries:
        ans = None
        while ans not in ("1", "2"):
            print(f"\n    {target_name}  [{report_name}]  ({reason})")
            print("      1. RETRY (checked manually, it exists)")
            print("      2. DROP (does not exist for this report)")
            ans = input("    Choice [1/2]: ").strip()
        decisions[(target_name, report_name)] = "retry" if ans == "1" else "drop"
    return decisions


def prompt_group_b(entries: list) -> dict:
    """
    entries: list of (target_name, report_name, conflicts) for Group B pairs.
    Returns {(target_name, report_name): "send"|"drop"}.
    """
    decisions = {}
    if not entries:
        return decisions
    print("\n  Group B -- found, but name conflict(s) detected:")
    for target_name, report_name, conflicts in entries:
        ans = None
        while ans not in ("1", "2"):
            print(f"\n    {target_name}  [{report_name}]  (also seen: {', '.join(conflicts)})")
            print("      1. GO AHEAD, SEND MAIL")
            print("      2. DROP (this report only -- other reports for this person are unaffected)")
            ans = input("    Choice [1/2]: ").strip()
        decisions[(target_name, report_name)] = "send" if ans == "1" else "drop"
    return decisions


def send_final(mailer: "Mailer", targets_by_name: dict, pair_state: dict,
               reports_by_target: dict, today: str,
               sent_names: list, failed_names: list, dropped_pairs: list,
               manual_pairs: list) -> None:
    """
    For each target in `reports_by_target`, send whatever of its reports
    are currently included (has a PDF, not excluded). Zero included
    reports -> no email, recorded as failed/manual instead.

    `pair_state`: {(target_name, report_name): {"pdf":, "status":,
    "conflicts":, "excluded": bool, "manual": bool}}. "manual" is only
    ever True for a Group A pair that was RETRIED and still didn't
    resolve -- a deliberate DROP (Group A "does not exist" or Group B
    "drop this report") is a resolved decision, not something needing
    follow-up, so it's reported separately from a genuine unresolved
    failure.
    """
    for target_name, report_names in reports_by_target.items():
        included_pdfs = []
        status_lines  = []
        for report_name in report_names:
            st = pair_state.get((target_name, report_name))
            if not st:
                continue
            if st["pdf"] and not st["excluded"]:
                included_pdfs.append(st["pdf"])
                status_lines.append(f"{report_name}: SENT")
            elif st["pdf"] and st["excluded"]:
                status_lines.append(f"{report_name}: DROPPED (by choice)")
                dropped_pairs.append((target_name, report_name))
            elif st.get("manual"):
                status_lines.append(f"{report_name}: FAILED (manual check needed)")
                manual_pairs.append((target_name, report_name))
            else:
                status_lines.append(f"{report_name}: DROPPED (confirmed not applicable)")
                dropped_pairs.append((target_name, report_name))

        log.info(f"  {target_name}: " + "; ".join(status_lines))

        if not included_pdfs:
            log.error(f"  All exports failed/dropped -- skipping email for {target_name}")
            failed_names.append(target_name)
            continue

        target = targets_by_name[target_name]
        try:
            subject = EMAIL_SUBJECT.format(aom_name=target_name, date=today)
            body    = EMAIL_BODY.format(aom_name=target_name, date=today)
            mailer.send(target["delivery_email"], subject, body, attachments=included_pdfs)
            log.info(
                f"  Email sent to {target['delivery_email']} with "
                f"{len(included_pdfs)}/{len(report_names)} attachment(s)"
            )
            sent_names.append(target_name)
        except Exception as e:
            log.error(f"  Email failed for {target_name}: {e}")
            failed_names.append(target_name)


def print_final_summary(group_name: str, total: int, sent: list, failed: list,
                         dropped_pairs: list, manual_pairs: list) -> None:
    """Print the consolidated final summary (actual email dispatch) for one group."""
    w = 52
    log.info("")
    log.info("=" * w)
    log.info(f"FINAL SUMMARY -- {group_name}")
    log.info("=" * w)
    log.info(f"Total targets        : {total}")
    log.info(f"Emails sent           : {len(sent)}")
    log.info(f"No email at all       : {len(failed)}")
    log.info(f"Reports dropped       : {len(dropped_pairs)}")
    log.info(f"Reports needing manual: {len(manual_pairs)}")
    if failed:
        log.info("No email was sent at all for these -- please handle manually:")
        for name in failed:
            log.info(f"  -> {name}")
    if manual_pairs:
        log.info("Retried and still unresolved -- these specific reports need a manual check")
        log.info("(the rest of that person's mail, if any, was still sent):")
        for name, report in manual_pairs:
            log.info(f"  -> {name} / {report}")
    if dropped_pairs:
        log.info("Resolved by decision (confirmed not applicable, or deliberately dropped --")
        log.info("no follow-up needed; the rest of that person's mail, if any, was still sent):")
        for name, report in dropped_pairs:
            log.info(f"  -> {name} / {report}")
    log.info("=" * w)
    log.info("")


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

    sent_by_group    = {g["group_name"]: [] for g in TARGET_GROUPS}
    failed_by_group  = {g["group_name"]: [] for g in TARGET_GROUPS}
    dropped_by_group = {g["group_name"]: [] for g in TARGET_GROUPS}
    manual_by_group  = {g["group_name"]: [] for g in TARGET_GROUPS}

    with PowerBIExporter(pbi_email, pbi_password) as exporter:

        for group in TARGET_GROUPS:
            gname   = group["group_name"]
            targets = group_targets[gname]
            reports = group["reports"]
            if not targets:
                log.info(f"\n  Group '{gname}': 0 targets loaded -- skipping.")
                continue

            log.info("\n" + "=" * 55)
            log.info(f"  PHASE 1 [{gname}] -- Processing all targets")
            log.info("=" * 55)

            reports_by_target = {t["name"]: [r["name"] for r in reports] for t in targets}
            all_pairs   = [(t, r) for t in targets for r in reports]
            pair_state  = {}   # (target_name, report_name) -> {pdf, status, conflicts, excluded}

            result = gather_for_pairs(exporter, group, all_pairs, targets, target_date, "PHASE 1")
            for key, info in result.items():
                pair_state[key] = {**info, "excluded": (info["status"] != "success"), "manual": False}

            rows = [
                (tname, rname,
                 st["status"] + (f" (conflicts: {st['conflicts']})" if st["conflicts"] else ""))
                for (tname, rname), st in sorted(pair_state.items())
            ]
            print_status_table(f"PHASE 1 RESULTS [{gname}] (per target / report)", rows)

            # Targets fully clean across every report -- send now, nothing to review.
            clean_targets = [
                tname for tname, rnames in reports_by_target.items()
                if all(pair_state[(tname, rn)]["status"] == "success" for rn in rnames)
            ]
            if clean_targets:
                send_final(
                    mailer, group_targets_by_name[gname], pair_state,
                    {t: reports_by_target[t] for t in clean_targets},
                    target_date, sent_by_group[gname], failed_by_group[gname],
                    dropped_by_group[gname], manual_by_group[gname]
                )

            review_targets = [t for t in reports_by_target if t not in clean_targets]
            if not review_targets:
                log.info(f"  Group '{gname}': Phase 1 achieved 100% clean success -- nothing to review.")
                continue

            # -- Build Group A / Group B entries from whatever isn't clean --
            group_a_entries = []
            group_b_entries = []
            for tname in review_targets:
                for rname in reports_by_target[tname]:
                    st = pair_state[(tname, rname)]
                    if st["status"] == "success":
                        continue
                    cls = classify_pair(st["status"], st["conflicts"])
                    if cls == "A":
                        if st["status"] == "conflict":
                            reason = (
                                f"{len(st['conflicts'])} conflicting names "
                                f"({', '.join(st['conflicts'])}) -- too many to be a plausible coincidence"
                            )
                        elif st["status"] == "dual":
                            reason = "unstable -- found on one attempt, not on the other"
                        else:
                            reason = "not found in the slicer"
                        group_a_entries.append((tname, rname, reason))
                    else:
                        group_b_entries.append((tname, rname, st["conflicts"]))

            print(
                f"\n  [{gname}] Review needed: {len(group_a_entries)} Group A pair(s), "
                f"{len(group_b_entries)} Group B pair(s)."
            )
            a_decisions = prompt_group_a(group_a_entries)
            b_decisions = prompt_group_b(group_b_entries)

            for key, decision in b_decisions.items():
                pair_state[key]["excluded"] = (decision == "drop")
            for key, decision in a_decisions.items():
                if decision == "drop":
                    pair_state[key]["excluded"] = True

            retry_keys = [key for key, d in a_decisions.items() if d == "retry"]

            if retry_keys:
                log.info("\n" + "=" * 55)
                log.info(f"  PHASE 2 [{gname}] -- Retrying user-confirmed pairs")
                log.info("=" * 55)
                log.info("  Re-logging in for a fresh session before Phase 2...")
                exporter.relogin()

                reports_by_name = {r["name"]: r for r in reports}
                targets_by_name = group_targets_by_name[gname]
                retry_pairs = [
                    (targets_by_name[tname], reports_by_name[rname])
                    for (tname, rname) in retry_keys
                ]
                retry_result = gather_for_pairs(exporter, group, retry_pairs, targets, target_date, "PHASE 2")

                post_retry_b_entries = []
                for key, info in retry_result.items():
                    cls = classify_pair(info["status"], info["conflicts"])
                    pair_state[key] = {
                        **info,
                        "excluded": (info["status"] != "success"),
                        # Only a still-Group-A retry outcome counts as a genuine
                        # unresolved failure needing manual follow-up -- success
                        # is included automatically, and Group B gets one more
                        # go-ahead/drop decision below (a resolved choice either way).
                        "manual": (cls == "A"),
                    }
                    if cls == "B":
                        post_retry_b_entries.append((key[0], key[1], info["conflicts"]))
                    # cls == "success" -> included automatically (excluded already False).
                    # cls == "A" again -> excluded=True, manual=True, final, no more retry offered.

                rows2 = [
                    (tname, rname,
                     info["status"] + (f" (conflicts: {info['conflicts']})" if info["conflicts"] else ""))
                    for (tname, rname), info in sorted(retry_result.items())
                ]
                print_status_table(f"PHASE 2 RETRY RESULTS [{gname}]", rows2)

                if post_retry_b_entries:
                    print(
                        f"\n  [{gname}] {len(post_retry_b_entries)} retried pair(s) now show a "
                        f"Group B conflict instead:"
                    )
                    post_decisions = prompt_group_b(post_retry_b_entries)
                    for key, decision in post_decisions.items():
                        pair_state[key]["excluded"] = (decision == "drop")

            # -- Final send for every target that had anything held back --
            send_final(
                mailer, group_targets_by_name[gname], pair_state,
                {t: reports_by_target[t] for t in review_targets},
                target_date, sent_by_group[gname], failed_by_group[gname],
                dropped_by_group[gname], manual_by_group[gname]
            )

    # -- Cleanup + final summary (one block per group) --
    cleanup_pdfs()
    for group in TARGET_GROUPS:
        gname = group["group_name"]
        targets = group_targets[gname]
        if not targets:
            continue
        print_final_summary(
            gname, len(targets), sent_by_group[gname], failed_by_group[gname],
            dropped_by_group[gname], manual_by_group[gname]
        )


if __name__ == "__main__":
    main()
