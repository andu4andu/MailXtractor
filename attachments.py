import logging
from concurrent.futures import ThreadPoolExecutor
from parsers import parse_single_attachment

logger = logging.getLogger(__name__)


def handle_attachments(email_struct: dict) -> list:
    return email_struct.get("attachments", [])


def spawn_attachment_workers(attachments: list) -> list:
    if not attachments:
        return []
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [
            executor.submit(parse_single_attachment, att["path"], att["filename"])
            for att in attachments
        ]
        results = []
        for att, future in zip(attachments, futures):
            try:
                results.append(future.result(timeout=300))
            except Exception as e:
                logger.error("Failed to parse '%s': %s", att['filename'], e)
                results.append((att["filename"], ""))
    return results


def combine_attachment_text(results: list) -> str:
    chunks = []
    for filename, text_chunk in results:
        if text_chunk.strip():
            chunks.append(f"[{filename}]\n{text_chunk}")
    return "\n\n".join(chunks)
