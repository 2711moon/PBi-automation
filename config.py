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
# List of reports to export and attach to the email for each AOM.
REPORTS = [
    {
        "workspace":     "Franchise Operations",
        "name":          "EBO Report -8 Day Wise Sale",
        "url":           "https://app.powerbi.com/groups/81d975dc-05d1-4d4e-b805-f5886368e433/reports/ac85b822-51b2-46af-b84c-7b0a2f5df834/f23a123c62d0a018a02e",
        "filter_table":  "Store Master",
        "filter_column": "AOM",
    }
    # To add more reports, simply copy the dictionary above:
    # {
    #     "workspace":     "Another Workspace",
    #     "name":          "Another Report",
    #     "filter_table":  "Store Master",
    #     "filter_column": "AOM Mail Id",
    # }
]

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
