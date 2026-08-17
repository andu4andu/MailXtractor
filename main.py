import json
import logging
import os
import re
import time
import datetime
from concurrent.futures import ThreadPoolExecutor
from readers import read_raw_email, extract_body_text
from attachments import handle_attachments, spawn_attachment_workers, combine_attachment_text
from extractor import load_rules, apply_extraction_rules
from reporter import generate_excel_report

logger = logging.getLogger(__name__)

with open("config.json", encoding="utf-8") as _f:
    CONFIG = json.load(_f)

INPUT_PATH = os.environ.get("MAILXTRACTOR_INPUT", CONFIG["input_path"])
_OUTPUT_DIR = os.environ.get("MAILXTRACTOR_OUTPUT", CONFIG["output_dir"])
INTERNAL_DOMAINS = tuple(CONFIG["internal_domains"])
RULES_PATH = "rules.json"


def get_output_path():
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(_OUTPUT_DIR, f"gold_report_{ts}.xlsx")


_RE_CTRL   = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_RE_FWD    = re.compile(r"^(?:(?:Fw|FW|Re|RE):\s*)+")
_RE_SEP    = re.compile(r"\s*---.*$")
_RE_SAMPLE = re.compile(r"\s*-\s*Sample\s*\d+\s*$", re.IGNORECASE)
_RE_CODE   = re.compile(r"\s*/\s*[A-Z0-9][A-Z0-9\-]+\s*$")
_RE_SUFFIX = re.compile(r"\s*-\s*[A-Za-z][A-Za-z0-9]+\s*$")


def _clean_subject(subject: str) -> str:
    subject = _RE_CTRL.sub("", subject).strip()
    subject = _RE_FWD.sub("", subject).strip()
    subject = _RE_SEP.sub("", subject).strip()
    subject = _RE_SAMPLE.sub("", subject).strip()
    subject = _RE_CODE.sub("", subject).strip()
    subject = _RE_SUFFIX.sub("", subject).strip()
    return subject


def _setup_logging() -> None:
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)-8s] %(name)s - %(message)s",
        datefmt="%H:%M:%S"
    )
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    fh = logging.FileHandler("run_log.txt", encoding="utf-8", mode="w")
    fh.setLevel(logging.INFO)
    fh.setFormatter(fmt)
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    root.addHandler(fh)
    root.addHandler(ch)


def process_email(folder_path: str, rules: dict) -> dict:
    folder_name = os.path.basename(folder_path)
    logger.info("Processing: %s", folder_name)

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
        logger.error("Failed to process %s: %s", folder_name, e)
        return {"_error": str(e), "folder": folder_name}


def main():
    rules = load_rules(RULES_PATH)
    records = []

    folders = [
        os.path.join(INPUT_PATH, name)
        for name in sorted(os.listdir(INPUT_PATH))
        if os.path.isdir(os.path.join(INPUT_PATH, name))
    ]

    logger.info("Found %d email folders", len(folders))

    t_start = time.time()
    with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
        futures = [(fp, executor.submit(process_email, fp, rules)) for fp in folders]
        for folder_path, future in futures:
            try:
                records.append(future.result(timeout=300))
                logger.info("Done: %s", os.path.basename(folder_path))
            except TimeoutError:
                logger.error("Timed out processing %s — skipping", os.path.basename(folder_path))
                records.append({"_error": "timeout", "folder": os.path.basename(folder_path)})

    logger.info("Total: %.1fs for %d folders", time.time() - t_start, len(folders))
    generate_excel_report(records, get_output_path())


if __name__ == "__main__":
    _setup_logging()
    main()
