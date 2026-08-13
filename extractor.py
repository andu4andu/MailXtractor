import json
import re
import uuid
from datetime import date


def load_rules(rules_path: str) -> dict:
    with open(rules_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _extract_label(text: str, labels: list, pattern: str = None, max_length: int = 100, line_start: bool = False, next_line: bool = False, collect_lines_pattern: str = None) -> str:
    for label in labels:
        escaped = re.escape(label)
        prefix = r"(?:^|\n)" if line_start else r""

        if collect_lines_pattern:
            for lm in re.finditer(rf"{prefix}{escaped}", text, re.IGNORECASE | re.MULTILINE):
                rest = text[lm.end():]
                lines = rest.split("\n")
                collected = []
                for i, line in enumerate(lines):
                    stripped = line.strip()
                    if i == 0:
                        stripped = re.sub(r'^[:\|\s]+', '', stripped)
                    if re.match(collect_lines_pattern, stripped, re.IGNORECASE):
                        cleaned = re.sub(r'[\s|]+$', '', stripped)
                        content = re.sub(r'^\d+\)\s*', '', cleaned).strip()
                        if content:
                            collected.append(cleaned)
                    elif collected:
                        break
                if collected:
                    return "; ".join(collected)
            continue

        match = re.search(
            rf"{prefix}{escaped}[^|\n]*?[:\|][^\S\n]*([^\n]+)",
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
                cleaned = re.sub(r'\s*\([A-Za-z]\)\s*|(?<=\d)(?:mm|cm|in)\b', ' ', value)
                cleaned = ' '.join(cleaned.split())
                pm = re.search(pattern, cleaned)
                return pm.group(0).strip() if pm else ""
            return value[:max_length].strip()

        if next_line and pattern:
            lm = re.search(rf"{prefix}{escaped}", text, re.IGNORECASE | re.MULTILINE)
            if lm:
                rest = text[lm.end():]
                lines_after = rest.split("\n")[1:]
                for line in lines_after[:5]:
                    if line.strip():
                        matches = re.findall(pattern, line, re.IGNORECASE)
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


def _keyword_match(text: str, values: list, aliases: dict = None) -> str:
    for value in values:
        match = re.search(rf"\b{re.escape(value)}\b", text, re.IGNORECASE)
        if match:
            return value
    if aliases:
        for alias, canonical in aliases.items():
            if re.search(rf"\b{re.escape(alias)}\b", text, re.IGNORECASE):
                return canonical
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


def _find_year_table_sum(text: str, max_value: int = None) -> str:
    year_pat = re.compile(r'(?<!\d)20[2-3]\d(?!\d)')
    num_pat = re.compile(r'[1-9]\d{1,2}(?:,\d{3})+|[1-9]\d{4,6}')
    lines = text.split('\n')

    def _collect(nums_found):
        values = []
        for n in nums_found:
            try:
                v = float(n.replace(',', ''))
                if max_value is None or v <= max_value:
                    values.append(v)
            except ValueError:
                pass
        if values and len(values) > 1 and int(values[-1]) == int(sum(values[:-1])):
            values = values[:-1]
        n = len(values)
        if n > 1 and n % 2 == 0:
            half1 = [int(v) for v in values[:n//2]]
            half2 = [int(v) for v in values[n//2:]]
            if half1 == half2:
                values = values[:n//2]
        return values

    # Approach 1: column-per-year — row with 2+ year tokens, quantities on next lines
    for i, line in enumerate(lines):
        if len(year_pat.findall(line)) >= 2:
            for j in range(i + 1, min(i + 6, len(lines))):
                candidate = lines[j]
                if len(year_pat.findall(candidate)) >= 2:
                    break
                values = _collect(num_pat.findall(candidate))
                if values:
                    return str(int(sum(values)))

    # Approach 2: row-per-year — year and quantity on the same line
    row_pat = re.compile(r'(?<!\d)20[2-3]\d(?![\d-])[^\n]*?\|\s*([1-9]\d{1,2}(?:,\d{3})+|[1-9]\d{4,6})')
    values = _collect(row_pat.findall(text))
    if values:
        return str(int(sum(values)))

    return ""


def _extract_sum_pattern(text: str, pattern: str, max_value: int = None, fallback_pattern: str = None) -> str:
    def _find_sum(pat):
        matches = re.findall(pat, text, re.IGNORECASE)
        values = []
        for m in matches:
            try:
                v = float(m.replace(",", ""))
                if max_value is None or v <= max_value:
                    values.append(v)
            except ValueError:
                pass
        return str(int(sum(values))) if values else ""

    result = _find_sum(pattern)
    if not result and fallback_pattern:
        result = _find_sum(fallback_pattern)
    return result


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


def _extract_column_values(text: str, labels: list) -> str:
    for label in labels:
        escaped = re.escape(label)
        for m in re.finditer(rf'\b{escaped}\b', text, re.IGNORECASE):
            line_start = text.rfind('\n', 0, m.start()) + 1
            line_end = text.find('\n', m.end())
            if line_end == -1:
                line_end = len(text)
            header_line = text[line_start:line_end]
            cols = [c.strip() for c in header_line.split(' | ')]
            col_idx = next((i for i, c in enumerate(cols) if re.search(rf'\b{escaped}\b', c, re.IGNORECASE)), None)
            if col_idx is None:
                continue
            rest = text[line_end + 1:]
            values = []
            for row_line in rest.split('\n'):
                if not row_line.strip() or row_line.startswith('['):
                    break
                row_cols = [c.strip() for c in row_line.split(' | ')]
                if col_idx < len(row_cols):
                    val = row_cols[col_idx].strip()
                    if val and re.search(r'[a-zA-Z]', val):
                        values.append(val)
            if values:
                return '; '.join(values)
    return ''


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
                value = _extract_label(text, rule.get("labels", []), rule.get("pattern"), line_start=rule.get("line_start", False), next_line=rule.get("next_line", False), collect_lines_pattern=rule.get("collect_lines_pattern"))
                if value and rule.get("normalize_boolean"):
                    value = _normalize_boolean(value)
                if not value and rule.get("column_values"):
                    value = _extract_column_values(text, rule.get("labels", []))
                if not value and rule.get("regex_fallback") and not rule.get("regex_fallback_body_first"):
                    value = _extract_regex(text, rule.get("regex_fallback"))
            elif rule_type == "regex":
                value = _extract_regex(text, rule.get("pattern", ""))
            elif rule_type == "max_pattern":
                value = _extract_max_pattern(text, rule.get("pattern", ""), rule.get("max_value"), rule.get("fallback_pattern"))
                if not value and rule.get("body_fallback_pattern") and text == body_string:
                    value = _extract_max_pattern(body_string, rule.get("body_fallback_pattern"), rule.get("max_value"))
            elif rule_type == "sum_pattern":
                chunks = re.split(r'\n\n(?=\[[^\]]+\.(?:pdf|xlsx?|pptx?|zip|txt|docx?|csv|msg|eml)\]\n)', text) if rule.get("per_file") else [text]
                if rule.get("year_anchored"):
                    sums = [_find_year_table_sum(c, rule.get("max_value")) for c in chunks]
                else:
                    sums = [_extract_sum_pattern(c, rule.get("pattern", ""), rule.get("max_value"), rule.get("fallback_pattern")) for c in chunks]
                value = next((s for s in sums if s), "")
            elif rule_type == "keyword_match":
                value = _keyword_match(text, rule.get("values", []), rule.get("aliases"))
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

        if not value and rule_type == "label" and rule.get("regex_fallback") and rule.get("regex_fallback_body_first"):
            for rb_text in [body_string, attachment_string]:
                value = _extract_regex(rb_text, rule.get("regex_fallback"))
                if value:
                    break

        if not value:
            value = rule.get("default", "")
        extracted[field] = value

    return extracted
