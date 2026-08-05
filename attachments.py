from concurrent.futures import ThreadPoolExecutor
from parsers import parse_single_attachment


def handle_attachments(email_struct: dict) -> list:
    return email_struct.get("attachments", [])


def spawn_attachment_workers(attachments: list) -> list:
    if not attachments:
        return []
    with ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(parse_single_attachment, att["path"], att["filename"])
            for att in attachments
        ]
        results = [future.result() for future in futures]
    return results


def combine_attachment_text(results: list) -> str:
    chunks = []
    for filename, text_chunk in results:
        if text_chunk.strip():
            chunks.append(f"[{filename}]\n{text_chunk}")
    return "\n\n".join(chunks)
