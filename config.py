"""
config.py -- Configuration for Power BI Report Automation.

Power BI credentials (email + password) are entered at runtime via terminal.
This file contains no secrets.
"""
import os

# == Paths =====================================================================
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
STOREMASTER = os.path.join(BASE_DIR, "Store Master.xlsx")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
LOG_FILE    = os.path.join(BASE_DIR, "automation.log")

# == Power BI (no URL needed - script discovers it automatically) ==============
# List of "target groups" -- each group is a designation to filter/email by
# (AOM, Cluster Manager, ...). Each group has its own slicer label in the
# report, its own Store Master columns to read values/delivery-emails from,
# and its own list of reports to export + attach into one email per target.
#
# Per-report keys:
#   filter_column -- which Store Master column's value to type into the
#                    slicer (usually == the group's value_column, but can
#                    differ, e.g. "AOM Mail Id" for a report whose slicer
#                    searches by email instead of name).
#   date_from / date_to -- "yesterday" | "month_start" | "today" |
#                    a literal "YYYY-MM-DD" date | None to skip setting the
#                    Date slicer entirely.
#   page          -- exact page name to export only that page, or None to
#                    export the whole report.
TARGET_GROUPS = [
    {
        "group_name":      "AOM",
        "slicer_label":    "AOM",
        "value_column":    "AOM",
        "delivery_column": "AutoEmail",
        "reports": [
            # Temporarily disabled -- testing Report 2's new date/page
            # features in isolation first, per user's plan. Re-enable once
            # confirmed working, then run with both reports together.
             {
                 "workspace":     "Franchise Operations",
                 "name":          "EBO Report -8 Day Wise Sale",
                 "url":           "https://app.powerbi.com/groups/81d975dc-05d1-4d4e-b805-f5886368e433/reports/63ea6b56-c09f-410d-a7ce-31381c174223/f23a123c62d0a018a02e",
                 "filter_column": "AOM",
                 "date_from":     None,
                 "date_to":       None,
                 "page":          None,
             },
            {
                "workspace":     "Franchise Operations",
                "name":          "EBO Report -5 DSR",
                "url":           "https://app.powerbi.com/groups/81d975dc-05d1-4d4e-b805-f5886368e433/reports/243700e6-fe17-4947-979e-f930fb07bb3e/ce3648bb53e8b9a3ea91",
                "filter_column": "AOM",
                "date_from":     "2026-09-01", #Expects the input in the form of year, month, and date. It will output in the format PowerBI currently expects, which is month, date, and year. 
                "date_to":       "2026-09-30",
                "page":          "Overall KPI Report",
            },
        ],
    },
    # To add another designation (e.g. Cluster Manager), copy the group above
    # and point value_column/delivery_column at the right Store Master
    # columns. NOTE: Store Master currently has "Cluster Manager" and
    # "Cluster Manager Mail Id" but no dedicated delivery-email column for
    # Cluster Managers yet (only a single "AutoEmail" column, which is
    # per-store/per-AOM) -- add that column to Store Master before
    # activating a Cluster Manager group, e.g.:
    # {
    #     "group_name":      "Cluster Manager",
    #     "slicer_label":    "Cluster Manager",
    #     "value_column":    "Cluster Manager",
    #     "delivery_column": "Cluster Manager AutoEmail",   # <- doesn't exist yet
    #     "reports": [ ... ],
    # },
]

# == Zoho SMTP =================================================================
SMTP_HOST         = "smtp.zoho.in"
SMTP_PORT         = 587
SMTP_USER         = "pragati.panhale@kisna.com"
SMTP_APP_PASSWORD = "JxQDD6QrPiMp"   # Zoho App Password Key

# == Email Templates ===========================================================
EMAIL_SUBJECT = "Franchise Report -- {aom_name} -- {date}"

EMAIL_BODY = """Dear {aom_name},

Please find attached your franchise performance report for {date}.

This is an automated report generated from Power BI.
Please do not reply to this email.
For any queries, reach out to the MIS team.

Regards,
Kisna Reporting Team"""

# == Send Time =================================================================
# The automation waits until this time before logging in and sending emails.
# Format: "HH:MM" in 24-hour time (e.g. "09:00" for 9 AM, "14:30" for 2:30 PM)
# To test immediately without waiting: run  python main.py --now
SEND_AT = "09:00"
