import email
import extract_msg
import re
import os
from html.parser import HTMLParser


EMAIL_EXTENSIONS = {".eml", ".msg"}


def read_raw_email(folder_path: str) -> dict:
    print(f"\n--- Processing folder: {os.path.basename(folder_path)} ---")

    email_file = None
    attachments = []

    for filename in sorted(os.listdir(folder_path)):
        file_path = os.path.join(folder_path, filename)
        ext = os.path.splitext(filename)[-1].lower()
        if ext in EMAIL_EXTENSIONS:
            if email_file is None:
                email_file = (file_path, ext)
                print(f"  [EMAIL FILE] {filename}")
            else:
                print(f"  [EMAIL FILE] skipping duplicate: {filename}")
        else:
            attachments.append({
                "filename": filename,
                "file_type": ext.lstrip("."),
                "path": file_path,
            })
            print(f"  [ATTACHMENT] {filename}")

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
            for part in msg.walk():
                content_disposition = part.get("Content-Disposition", "")
                if part.get_content_type() == "text/plain" and "attachment" not in content_disposition:
                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                    break

        elif ext == ".msg":
            msg = extract_msg.Message(file_path)
            subject = msg.subject or ""
            sender = msg.sender or ""
            recipients = msg.to or ""
            date = str(msg.date) if msg.date else ""
            html_body = msg.htmlBody or b""
            if html_body.strip():
                body = html_body.decode("utf-8", errors="ignore")
            else:
                body = msg.body or ""
            print(f"  [MSG HTML BODY] {repr(body[:100])}")

        print(f"  [SUBJECT]  {subject}")
        print(f"  [SENDER]   {sender}")
        print(f"  [DATE]     {date}")
        print(f"  [BODY]     {body[:100].strip()}{'...' if len(body) > 100 else ''}")
    else:
        print(f"  [NO EMAIL FILE] attachments only")

    print(f"  [ATTACHMENTS COUNT] {len(attachments)}")

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
    print(f"  [BODY TEXT] {body[:100].strip()}{'...' if len(body) > 100 else ''}")
    return body
