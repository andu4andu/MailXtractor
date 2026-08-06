import json
import re
import uuid
from datetime import date


def load_rules(rules_path: str) -> dict:
    with open(rules_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _extract_label(text: str, labels: list, pattern: str = None, max_length: int = 100, line_start: bool = False) -> str:
    for label in labels:
        escaped = re.escape(label)
        prefix = r"(?:^|\n)" if line_start else r""
        match = re.search(
            rf"{prefix}{escaped}[^|\n]*[:\|]\s*([^|\n]+)",
            text,
            re.IGNORECASE | re.MULTILINE
        )
        if match:
            value = match.group(1).strip()
            if pattern:
                pm = re.search(pattern, value)
                return pm.group(0).strip() if pm else ""
            return value[:max_length].strip()
    return ""


def _extract_regex(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(0).strip() if match else ""


def _extract_supplier_from_recipients(body: str) -> str:
    match = re.search(r"Recipients:\s*(.+)", body)
    if not match:
        return ""
    recipients_line = match.group(1)
    emails = re.findall(r"@([\w.\-]+)", recipients_line)
    for domain in emails:
        parts = domain.split(".")
        company = parts[0] if len(parts) >= 2 else domain
        if company.lower() not in ("gmail", "yahoo", "hotmail", "outlook"):
            return company.replace("-", " ").replace("_", " ").title()
    return ""


def _keyword_match(text: str, values: list, debug_field: str = "") -> str:
    for value in values:
        match = re.search(rf"\b{re.escape(value)}\b", text, re.IGNORECASE)
        if match:
            if debug_field:
                start = max(0, match.start() - 50)
                end = min(len(text), match.end() + 50)
                print(f"  [MATCH:{debug_field}] '{value}' found in: ...{repr(text[start:end])}...")
            return value
    return ""


def _normalize_boolean(value: str) -> str:
    if re.search(r"\b(yes|y)\b", value, re.IGNORECASE):
        return "Yes"
    if re.search(r"\b(no|n|n/a)\b", value, re.IGNORECASE):
        return "No"
    return value


def _extract_max_pattern(text: str, pattern: str) -> str:
    matches = re.findall(pattern, text, re.IGNORECASE)
    values = []
    for m in matches:
        try:
            values.append(float(m.replace(",", "")))
        except ValueError:
            pass
    return str(int(max(values))) if values else ""


def apply_extraction_rules(
    body_string: str,
    attachment_string: str,
    rules: dict,
    correlation_id: str = None
) -> dict:
    extracted = {}

    for field, rule in rules.items():
        rule_type = rule.get("type")
        source = rule.get("source_priority", "both")

        if rule_type == "auto":
            if field == "correlation_id":
                extracted[field] = correlation_id or str(uuid.uuid4())
            elif field == "date_of_input":
                extracted[field] = date.today().isoformat()
            continue

        if rule_type == "empty":
            extracted[field] = ""
            continue

        if source == "body":
            sources = [body_string]
        elif source == "attachment":
            sources = [attachment_string]
        else:
            sources = [attachment_string, body_string]

        value = ""
        for text in sources:
            if rule_type == "label":
                value = _extract_label(text, rule.get("labels", []), rule.get("pattern"), line_start=rule.get("line_start", False))
                if value and rule.get("normalize_boolean"):
                    value = _normalize_boolean(value)
            elif rule_type == "regex":
                value = _extract_regex(text, rule.get("pattern", ""))
            elif rule_type == "max_pattern":
                value = _extract_max_pattern(text, rule.get("pattern", ""))
            elif rule_type == "keyword_match":
                value = _keyword_match(text, rule.get("values", []), debug_field=field)
            elif rule_type == "email_domain":
                value = _extract_supplier_from_recipients(body_string)
                if not value and rule.get("fallback_type") == "label":
                    value = _extract_label(text, rule.get("labels", []), rule.get("pattern"))
            if value:
                break

        if not value:
            value = rule.get("default", "")
        extracted[field] = value

    return extracted
