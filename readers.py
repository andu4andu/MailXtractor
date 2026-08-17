import email
import extract_msg
import logging
import re
import os
from html.parser import HTMLParser

logger = logging.getLogger(__name__)
logging.getLogger("extract_msg").setLevel(logging.ERROR)

EMAIL_EXTENSIONS = {".eml", ".msg"}


def read_raw_email(folder_path: str) -> dict:
    logger.info("Processing folder: %s", os.path.basename(folder_path))

    email_file = None
    attachments = []

    for filename in sorted(os.listdir(folder_path)):
        file_path = os.path.join(folder_path, filename)
        ext = os.path.splitext(filename)[-1].lower()
        if ext in EMAIL_EXTENSIONS:
            if email_file is None:
                email_file = (file_path, ext)
                logger.debug("Email file: %s", filename)
            else:
                logger.warning("Skipping duplicate email file: %s", filename)
        else:
            attachments.append({
                "filename": filename,
                "file_type": ext.lstrip("."),
                "path": file_path,
            })
            logger.debug("Attachment: %s", filename)

    subject, sender, recipients, date, body = "", "", "", "", ""

    if email_file:
        file_path, ext = email_file
        if ext == ".eml":
            with open(file_path, "rb") as f:
                msg = email.message_from_bytes(f.read())
            subject = msg.get("Subject", "")
            sender = msg.get("From", "")
            recipients = msg.get("To", "")
            date = msg.get("Date", "")
            html_fallback = ""
            for part in msg.walk():
                content_disposition = part.get("Content-Disposition", "")
                if "attachment" in content_disposition:
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                if part.get_content_type() == "text/plain":
                    body = payload.decode("utf-8", errors="ignore")
                    if body:
                        break
                if part.get_content_type() == "text/html" and not html_fallback:
                    html_fallback = payload.decode("utf-8", errors="ignore")
            if not body and html_fallback:
                body = html_fallback

        elif ext == ".msg":
            with extract_msg.Message(file_path) as msg:
                subject = msg.subject or ""
                sender = msg.sender or ""
                recipients = msg.to or ""
                date = str(msg.date) if msg.date else ""
                html_body = msg.htmlBody or b""
                if html_body.strip():
                    body = html_body.decode("utf-8", errors="ignore")
                else:
                    body = msg.body or ""
            logger.debug("MSG body preview: %s", repr(body[:100]))

        logger.debug("Subject: %s", subject)
        logger.debug("Sender:  %s", sender)
        logger.debug("Date:    %s", date)
        logger.debug("Body:    %.100s%s", body.strip(), "..." if len(body) > 100 else "")
    else:
        logger.warning("No email file found in %s — attachments only", os.path.basename(folder_path))

    logger.debug("Attachments: %d", len(attachments))

    return {
        "subject": subject,
        "sender": sender,
        "recipients": recipients,
        "date": date,
        "body": body,
        "attachments": attachments,
    }


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self):
        return " ".join(self._parts)


def extract_body_text(email_struct: dict) -> str:
    body = email_struct.get("body", "")
    stripper = _HTMLStripper()
    stripper.feed(body)
    body = re.sub(r"\s+", " ", stripper.get_text()).strip()
    logger.debug("Body text: %.100s%s", body.strip(), "..." if len(body) > 100 else "")
    return body
