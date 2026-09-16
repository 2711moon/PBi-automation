"""
mailer.py - Sends emails via Zoho SMTP with a PDF attachment.
"""
import smtplib
import os
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from config import SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_APP_PASSWORD

log = logging.getLogger(__name__)


class Mailer:
    def send(self, to_email: str, subject: str, body: str, attachment_path: str) -> None:
        """Send one email with a single PDF attachment via Zoho SMTP."""
        msg = MIMEMultipart()
        msg["From"]    = SMTP_USER
        msg["To"]      = to_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain", "utf-8"))

        filename = os.path.basename(attachment_path)
        with open(attachment_path, "rb") as fh:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(fh.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_APP_PASSWORD)
            server.sendmail(SMTP_USER, to_email, msg.as_string())

        log.info(f"  Email sent -> {to_email}")
