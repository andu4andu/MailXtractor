import json
import os
import re
import time
import datetime
from concurrent.futures import ThreadPoolExecutor
from readers import read_raw_email, extract_body_text
from attachments import handle_attachments, spawn_attachment_workers, combine_attachment_text
from extractor import load_rules, apply_extraction_rules, INTERNAL_DOMAINS
from reporter import generate_excel_report

with open("config.json", encoding="utf-8") as _f:
    CONFIG = json.load(_f)

INPUT_PATH = CONFIG["input_path"]
INTERNAL_DOMAINS = tuple(CONFIG["internal_domains"])
RULES_PATH = "rules.json"


def get_output_path():
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(CONFIG["output_dir"], f"gold_report_{ts}.xlsx")


def _clean_subject(subject: str) -> str:
    subject = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", subject).strip()
    subject = re.sub(r"^(?:(?:Fw|FW|Re|RE):\s*)+", "", subject).strip()
    subject = re.sub(r"\s*---.*$", "", subject).strip()
    subject = re.sub(r"\s*-\s*Sample\s*\d+\s*$", "", subject, flags=re.IGNORECASE).strip()
    subject = re.sub(r"\s*/\s*[A-Z0-9][A-Z0-9\-]+\s*$", "", subject).strip()
    subject = re.sub(r"\s*-\s*[A-Za-z][A-Za-z0-9]+\s*$", "", subject).strip()
    return subject


def process_email(folder_path: str, rules: dict) -> dict:
    folder_name = os.path.basename(folder_path)
    print(f"\n{'='*60}")
    print(f"Processing: {folder_name}")
    print(f"{'='*60}")

    try:
        email_struct = read_raw_email(folder_path)
        body_string = extract_body_text(email_struct)
        recipients_raw = email_struct.get("recipients", "")
        external_recipients = " ".join(
            r for r in re.split(r"[,;]", recipients_raw)
            if not any(d in r.lower() for d in INTERNAL_DOMAINS)
        )
        cleaned_subject = _clean_subject(email_struct.get("subject", ""))
        email_header = "\n".join(filter(None, [
            f"Subject: {email_struct.get('subject', '')}",
            f"Project: {cleaned_subject}" if cleaned_subject else "",
            f"Sender: {email_struct.get('sender', '')}",
            f"Recipients: {external_recipients}",
            f"Date: {email_struct.get('date', '')}",
        ]))
        full_body = email_header + "\n" + body_string
        attachments = handle_attachments(email_struct)
        results = spawn_attachment_workers(attachments)
        attachment_string = combine_attachment_text(results)
        extracted = apply_extraction_rules(full_body, attachment_string, rules)
        extracted["correlation_id"] = folder_name
        return extracted

    except Exception as e:
        print(f"  [ERROR] {folder_name}: {e}")
        return {"_error": str(e), "folder": folder_name}


def main():
    rules = load_rules(RULES_PATH)
    records = []

    folders = [
        os.path.join(INPUT_PATH, name)
        for name in sorted(os.listdir(INPUT_PATH))
        if os.path.isdir(os.path.join(INPUT_PATH, name))
    ]

    print(f"Found {len(folders)} email folders")

    t_start = time.time()
    with ThreadPoolExecutor() as executor:
        futures = [(fp, executor.submit(process_email, fp, rules)) for fp in folders]
        for folder_path, future in futures:
            records.append(future.result())
            print(f"  [TIMER] {os.path.basename(folder_path)}")

    print(f"\n[TIMER] Total: {time.time() - t_start:.1f}s for {len(folders)} folders")
    generate_excel_report(records, get_output_path())


if __name__ == "__main__":
    import sys
    with open("run_log.txt", "w", encoding="utf-8") as log:
        sys.stdout = log
        main()
    sys.stdout = sys.__stdout__
    print("Done. Check run_log.txt and gold_report.xlsx")
