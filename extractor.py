import json
import re
import uuid
from datetime import date


def load_rules(rules_path: str) -> dict:
    with open(rules_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _extract_label(text: str, labels: list, pattern: str = None, max_length: int = 100, line_start: bool = False, next_line: bool = False) -> str:
    for label in labels:
        escaped = re.escape(label)
        prefix = r"(?:^|\n)" if line_start else r""
        match = re.search(
            rf"{prefix}{escaped}[^|\n]*?[:\|]\s*([^\n]+)",
            text,
            re.IGNORECASE | re.MULTILINE
        )
        if match:
            raw = match.group(1).strip()
            parts = [p.strip() for p in raw.split("|")]
            value = next((p for p in parts if p), "")
            if not value:
                continue
            if pattern:
                pm = re.search(pattern, value)
                return pm.group(0).strip() if pm else ""
            return value[:max_length].strip()

        if next_line and pattern:
            lm = re.search(rf"{prefix}{escaped}", text, re.IGNORECASE | re.MULTILINE)
            if lm:
                rest = text[lm.end():]
                parts = rest.split("\n", 2)
                next_line_text = parts[1] if len(parts) > 1 else ""
                if next_line_text:
                    matches = re.findall(pattern, next_line_text, re.IGNORECASE)
                    if matches:
                        return matches[-1].strip()
    return ""


def _extract_regex(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(0).strip() if match else ""


INTERNAL_DOMAINS = ("continental", "aumovio")


def _extract_supplier_from_body(body: str) -> str:
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", body)
    for domain in emails:
        if not any(d in domain.lower() for d in INTERNAL_DOMAINS):
            company = domain.split(".")[0]
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
    if re.match(r"^[\d\s.,]+$", value):
        return ""
    if value.strip():
        return "Yes"
    return ""


def _extract_max_pattern(text: str, pattern: str, max_value: int = None, fallback_pattern: str = None) -> str:
    def _find_max(pat):
        matches = re.findall(pat, text, re.IGNORECASE)
        values = []
        for m in matches:
            try:
                v = float(m.replace(",", ""))
                if max_value is None or v <= max_value:
                    values.append(v)
            except ValueError:
                pass
        return str(int(max(values))) if values else ""

    result = _find_max(pattern)
    if not result and fallback_pattern:
        result = _find_max(fallback_pattern)
    return result


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
                value = _extract_label(text, rule.get("labels", []), rule.get("pattern"), line_start=rule.get("line_start", False), next_line=rule.get("next_line", False))
                if value and rule.get("normalize_boolean"):
                    value = _normalize_boolean(value)
            elif rule_type == "regex":
                value = _extract_regex(text, rule.get("pattern", ""))
            elif rule_type == "max_pattern":
                value = _extract_max_pattern(text, rule.get("pattern", ""), rule.get("max_value"), rule.get("fallback_pattern"))
                if not value and rule.get("body_fallback_pattern") and text == body_string:
                    value = _extract_max_pattern(body_string, rule.get("body_fallback_pattern"), rule.get("max_value"))
            elif rule_type == "keyword_match":
                value = _keyword_match(text, rule.get("values", []), debug_field=field)
            elif rule_type == "email_domain":
                value = _extract_supplier_from_body(body_string)
                if not value and rule.get("fallback_type") == "label":
                    for fallback_text in [attachment_string, body_string]:
                        value = _extract_label(fallback_text, rule.get("labels", []), rule.get("pattern"))
                        if value:
                            break
                if value and any(d in value.lower() for d in ("continental", "aumovio")):
                    value = ""
            if value:
                break

        if not value:
            value = rule.get("default", "")
        extracted[field] = value

    return extracted
