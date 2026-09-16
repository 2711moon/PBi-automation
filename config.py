"""
config.py -- Configuration for Power BI Report Automation.

Power BI credentials (email + password) are entered at runtime via terminal.
This file contains no secrets.
"""
import os

# == Paths =====================================================================
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
STOREMASTER = os.path.join(BASE_DIR, "StoreMAster.xlsx")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")
LOG_FILE    = os.path.join(BASE_DIR, "automation.log")

# == Power BI REST API =========================================================
# IDs taken from the Power BI Service URL (confirmed in test runs).
PBI_GROUP_ID   = "df64deac-1436-446d-a5cf-f7f30103f7c5"   # Workspace (group) ID
PBI_REPORT_ID  = "6aacdb66-1554-4930-84da-7b40588ecb22"   # Report ID
PBI_RLS_ROLE   = "AOM"   # RLS role name exactly as defined in PBI Desktop → Manage roles

# Human-readable labels (logs only)
PBI_WORKSPACE_NAME = "AutoEmails"
PBI_REPORT_NAME    = "Franchise DataSet"

# == Zoho SMTP =================================================================
SMTP_HOST         = "smtp.zoho.in"
SMTP_PORT         = 587
SMTP_USER         = "pragati.panhale@kisna.com"
SMTP_APP_PASSWORD = "JxQDD6QrPiMp"   # Zoho App Password

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
