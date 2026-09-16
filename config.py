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

# == Power BI (no URL needed - script discovers it automatically) ==============
PBI_WORKSPACE_NAME = "AutoEmails"         # Exact workspace name in Power BI Service
PBI_REPORT_NAME    = "Franchise DataSet"  # Exact report name in that workspace

# Exact table + column in the Power BI data model used for RLS URL filtering
PBI_FILTER_TABLE  = "Store Master"
PBI_FILTER_COLUMN = "AOM Mail Id"

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
