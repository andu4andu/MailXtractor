import io
import os
import tempfile
import zipfile
import openpyxl
import pdfplumber
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pytesseract
from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\uik11822\OCR\tesseract.exe"

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


def _cell_str(cell) -> str:
    if isinstance(cell, float):
        if cell == int(cell):
            return str(int(cell))
        return f"{cell:.2f}"
    return str(cell)


def parse_xlsx(path: str) -> str:
    print(f"  [PARSE XLSX] {path}")
    wb = openpyxl.load_workbook(path, data_only=True)
    text = ""
    for sheet in wb.worksheets:
        if sheet.sheet_state in ("hidden", "veryHidden"):
            print(f"  [PARSE XLSX] skipping hidden sheet: {sheet.title}")
            continue
        if any(skip in sheet.title for skip in ["RM_Catalog", "RM_Attribute", "BExRepository", "Drop Downs", "Instructions", "Explanation"]):
            print(f"  [PARSE XLSX] skipping reference sheet: {sheet.title}")
            continue
        text += f"[Sheet: {sheet.title}]\n"
        for row in sheet.iter_rows(values_only=True):
            cells = [_cell_str(cell) for cell in row if cell is not None and _cell_str(cell).strip()]
            if cells:
                text += " | ".join(cells) + "\n"
    print(f"  [PARSE XLSX] extracted {len(text)} characters")
    return text.strip()


def _is_old_ppt_format(path: str) -> bool:
    with open(path, "rb") as f:
        return f.read(4) == b"\xD0\xCF\x11\xE0"


def _convert_ppt_to_pptx(path: str) -> str:
    import win32com.client
    abs_path = os.path.abspath(path)
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    temp_path = temp_file.name
    temp_file.close()
    os.remove(temp_path)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(abs_path, WithWindow=False)
    deck.SaveAs(temp_path, 24)
    deck.Close()
    powerpoint.Quit()
    print(f"  [CONVERT PPT] saved to {temp_path}, exists: {os.path.exists(temp_path)}")
    return temp_path


def _iter_shapes(shapes):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes)
        else:
            yield shape


def parse_ppt(path: str) -> str:
    print(f"  [PARSE PPT] {path}")
    converted_path = None
    try:
        if _is_old_ppt_format(path):
            print(f"  [PARSE PPT] old format detected, converting...")
            converted_path = _convert_ppt_to_pptx(path)
            pptx_path = converted_path
        else:
            pptx_path = path
        prs = Presentation(pptx_path)
    except Exception as e:
        print(f"  [PARSE PPT] failed to open: {e}")
        if converted_path and os.path.exists(converted_path):
            os.remove(converted_path)
        return ""
    text = ""
    for slide_num, slide in enumerate(prs.slides, start=1):
        text += f"[Slide {slide_num}]\n"
        for shape in _iter_shapes(slide.shapes):
            if shape.has_text_frame and not (shape.is_placeholder and shape.placeholder_format.idx == 12):
                for para in shape.text_frame.paragraphs:
                    line = para.text.strip()
                    if line:
                        text += line + "\n"
            if shape.has_table:
                for row in shape.table.rows:
                    row_text = " | ".join(cell.text.replace("\n", " ").strip() for cell in row.cells)
                    if row_text.strip("|").strip():
                        text += row_text + "\n"
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = Image.open(io.BytesIO(shape.image.blob))
                ocr_text = pytesseract.image_to_string(image).strip()
                if ocr_text:
                    print(f"  [PARSE PPT] OCR on image in slide {slide_num}")
                    text += ocr_text + "\n"
    if converted_path and os.path.exists(converted_path):
        os.remove(converted_path)
    print(f"  [PARSE PPT] extracted {len(text)} characters")
    return text.strip()


def parse_zip(path: str) -> str:
    print(f"  [PARSE ZIP] {path}")
    text = ""
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(path, "r") as z:
            z.extractall(temp_dir)
        for root, _, files in os.walk(temp_dir):
            for filename in files:
                file_path = os.path.join(root, filename)
                ext = os.path.splitext(filename)[-1].lower()
                print(f"  [PARSE ZIP] found {filename}")
                if ext in PDF_EXTENSIONS:
                    text += f"[{filename}]\n" + parse_pdf(file_path) + "\n"
                elif ext in SPREADSHEET_EXTENSIONS:
                    text += f"[{filename}]\n" + parse_xlsx(file_path) + "\n"
                elif ext in PRESENTATION_EXTENSIONS:
                    text += f"[{filename}]\n" + parse_ppt(file_path) + "\n"
                elif ext in DOCUMENT_EXTENSIONS:
                    text += f"[{filename}]\n" + parse_txt(file_path) + "\n"
                elif ext in ZIP_EXTENSIONS:
                    text += f"[{filename}]\n" + parse_zip(file_path) + "\n"
                else:
                    print(f"  [PARSE ZIP] skipping unsupported file: {filename}")
    finally:
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for f in files:
                os.remove(os.path.join(root, f))
            for d in dirs:
                os.rmdir(os.path.join(root, d))
        os.rmdir(temp_dir)
    print(f"  [PARSE ZIP] extracted {len(text)} characters")
    return text.strip()


def parse_txt(path: str) -> str:
    print(f"  [PARSE TXT] {path}")
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read().strip()
    print(f"  [PARSE TXT] extracted {len(text)} characters")
    return text


def parse_single_attachment(path: str, filename: str) -> tuple:
    ext = os.path.splitext(filename)[-1].lower()
    print(f"  [DISPATCHER] {filename} ({ext})")

    if ext in PDF_EXTENSIONS:
        text = parse_pdf(path)
    elif ext in SPREADSHEET_EXTENSIONS:
        text = parse_xlsx(path)
    elif ext in PRESENTATION_EXTENSIONS:
        text = parse_ppt(path)
    elif ext in DOCUMENT_EXTENSIONS:
        text = parse_txt(path)
    elif ext in ZIP_EXTENSIONS:
        text = parse_zip(path)
    elif ext in IMAGE_EXTENSIONS:
        image = Image.open(path)
        text = pytesseract.image_to_string(image).strip()
        print(f"  [DISPATCHER] OCR on image: {len(text)} characters")
    else:
        print(f"  [DISPATCHER] unsupported file type: {ext}, skipping")
        text = ""

    return (filename, text)
