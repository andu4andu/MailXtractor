import pdfplumber
import pytesseract
from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\uik11822\OCRtesseract.exe"

IMAGE_EXTENSIONS = {".png", ".jpeg", ".jpg", ".bmp", ".tiff"}
PDF_EXTENSIONS = {".pdf"}
SPREADSHEET_EXTENSIONS = {".xlsx", ".xls", ".csv"}
PRESENTATION_EXTENSIONS = {".ppt", ".pptx"}
DOCUMENT_EXTENSIONS = {".doc", ".docx", ".txt"}
ZIP_EXTENSIONS = {".zip", ".rar", ".7z"}



def _join_cell(cell: str) -> str:
    parts = cell.split("\n")
    result = parts[0]
    for part in parts[1:]:
        stripped = part.strip()
        is_fragment = (
            result and result[-1].isalpha() and
            stripped and stripped[0].islower() and
            stripped.isalpha() and
            len(stripped) <= 5
        )
        result += stripped if is_fragment else " " + stripped
    return result


def parse_pdf(path: str) -> str:
    print(f"  [PARSE PDF] {path}")
    text = ""

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            table_bboxes = [table.bbox for table in page.find_tables()]
            tables = page.extract_tables()
            table_text = ""
            for table in tables:
                for row in table:
                    row_text = " | ".join(_join_cell(cell) if cell else "" for cell in row)
                    table_text += row_text + "\n"

            non_table_page = page
            for bbox in table_bboxes:
                non_table_page = non_table_page.filter(lambda obj, bb=bbox: not (bb[0] <= obj["x0"] <= bb[2] and bb[1] <= obj["top"] <= bb[3]))
            words = non_table_page.extract_words(x_tolerance=5, y_tolerance=3)
            words = sorted(words, key=lambda w: (round(w["top"] / 5) * 5, w["x0"]))
            lines = {}
            for w in words:
                line_key = round(w["top"] / 5) * 5
                lines.setdefault(line_key, []).append(w["text"])
            body_text = "\n".join(" ".join(line) for line in lines.values())

            text += body_text + "\n" + table_text

    text = text.strip()

    if not text:
        print(f"  [PARSE PDF] no text found, falling back to OCR")
        images = convert_from_path(path)
        for image in images:
            text += pytesseract.image_to_string(image) + "\n"
        text = text.strip()

    print(f"  [PARSE PDF] extracted {len(text)} characters")
    return text


def parse_txt(path: str) -> str:
    print(f"  [PARSE TXT] {path}")
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read().strip()
    print(f"  [PARSE TXT] extracted {len(text)} characters")
    return text
