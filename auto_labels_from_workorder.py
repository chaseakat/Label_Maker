import os
import sys
import re
import zipfile
import pdfplumber
import pytesseract
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader


# ========================
# CONFIG
# ========================
LOGO_ZIP_PATH = None  # e.g. r"C:\path\to\jti logo.zip"
LOGO_IMAGE_PATH = r"C:\Users\rnicks\OneDrive - JTI Millwork\Desktop\Label_Maker\JTI_LOGO.jpg"

OUTPUT_DIR = r"J:\Labels"

PAGE_W, PAGE_H = landscape((4 * inch, 6 * inch))

BORDER_MARGIN = 10
INNER_MARGIN_X = 30
TOP_MARGIN_Y = 26
BOTTOM_MARGIN_Y = 26

JOB_FONT_SIZE = 26
ITEM_FONT_SIZE = 26
DESC_FONT_SIZE = 18

LOGO_SHRINK = 0.60
ITEM_DESC_GAP = 18


# ========================
# HELPERS
# ========================
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def sanitize_filename(text: str) -> str:
    text = (text or "").strip()
    text = re.sub(r'[<>:"/\\|?*]', "", text)
    text = re.sub(r"\s+", "_", text)
    return text

def unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while True:
        p = f"{base}_{i}{ext}"
        if not os.path.exists(p):
            return p
        i += 1

def pick_output_path(job_number: str, job_name: str) -> str:
    job_digits = (job_number or "").replace("#", "").strip()
    job_name_clean = (job_name or "").strip()

    if job_digits and job_name_clean.lower().startswith(job_digits.lower()):
        job_name_clean = job_name_clean[len(job_digits):].strip(" -_")

    safe_job = sanitize_filename(job_digits or "#")
    safe_name = sanitize_filename(job_name_clean or "UNKNOWN_JOB")

    filename = f"{safe_job}_{safe_name}_Labels.pdf"
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return unique_path(os.path.join(OUTPUT_DIR, filename))


# ========================
# LOGO
# ========================
def extract_logo_from_zip(zip_path: str) -> str:
    extract_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__logo_extract__")
    os.makedirs(extract_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(extract_dir)

    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.lower().endswith((".png", ".jpg", ".jpeg", ".jfif", ".webp", ".bmp")):
                return os.path.join(root, f)
    raise RuntimeError("No logo image found inside logo zip.")

def get_logo_path() -> str:
    if LOGO_IMAGE_PATH and os.path.exists(LOGO_IMAGE_PATH):
        return LOGO_IMAGE_PATH
    if LOGO_ZIP_PATH and os.path.exists(LOGO_ZIP_PATH):
        return extract_logo_from_zip(LOGO_ZIP_PATH)
    raise RuntimeError("Logo not found. Set LOGO_IMAGE_PATH or LOGO_ZIP_PATH.")


# ========================
# TEXT EXTRACTION
# ========================
def extract_text_from_pdf(path: str) -> str:
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text

def extract_text_from_image(path: str) -> str:
    return pytesseract.image_to_string(Image.open(path))


# ========================
# NO-DIMS CLEANER
# ========================
NUM_TOKEN = r"(?:\d+(?:\.\d+)?)"

W_DIM_BLOCK_RE = re.compile(
    rf"\bw\s+(?:{NUM_TOKEN}\s+){{2,8}}(?:L|R|Left|Right|Both|None)?\b",
    re.IGNORECASE
)

TRAIL_NUMS_RE = re.compile(
    rf"(?:\s+{NUM_TOKEN}){{1,10}}\s*(?:L|R|Left|Right|Both|None)?\s*$",
    re.IGNORECASE
)

FE_RE = re.compile(r"\bFE\s*[:=\-]?\s*[A-Za-z0-9]+\b", re.IGNORECASE)
ITEM_TOKEN_RE = re.compile(r"\bItem\s*\d+\.\d+\b", re.IGNORECASE)

def clean_description_no_dims(raw: str) -> str:
    s = _norm(raw)
    s = ITEM_TOKEN_RE.sub("", s)
    s = _norm(s)
    s = FE_RE.sub("", s)
    s = _norm(s)
    s = W_DIM_BLOCK_RE.sub("", s)
    s = _norm(s)
    s = TRAIL_NUMS_RE.sub("", s)
    s = _norm(s)
    s = re.sub(r"\s+(L|R|Left|Right|Both|None)\s*$", "", s, flags=re.IGNORECASE)
    return _norm(s)


# ========================
# TRUNCATION
# ========================
def truncate_text(c, text, font_name, font_size, max_width):
    if not text:
        return ""
    ellipsis = "..."
    if c.stringWidth(text, font_name, font_size) <= max_width:
        return text
    max_width = max(0, max_width - c.stringWidth(ellipsis, font_name, font_size))
    t = text
    while t and c.stringWidth(t, font_name, font_size) > max_width:
        t = t[:-1]
    return (t + ellipsis) if t else ellipsis


# ========================
# LABEL DRAWING
# ========================
def draw_label(c, logo_path: str, job_number: str, job_name: str, item_label: str, description: str):
    job_number = job_number or ""
    job_name = job_name or ""
    item_label = item_label or ""
    description = description or ""

    c.setStrokeColor(colors.black)
    c.setLineWidth(4)
    c.rect(BORDER_MARGIN, BORDER_MARGIN, PAGE_W - 2 * BORDER_MARGIN, PAGE_H - 2 * BORDER_MARGIN)

    top_y = PAGE_H - BORDER_MARGIN - TOP_MARGIN_Y

    c.setFont("Helvetica-Bold", JOB_FONT_SIZE)
    c.drawString(INNER_MARGIN_X, top_y, job_number)

    name_size = 24
    max_name_w = PAGE_W - (2 * INNER_MARGIN_X) - 120
    while name_size > 10 and c.stringWidth(job_name, "Helvetica", name_size) > max_name_w:
        name_size -= 1
    c.setFont("Helvetica", name_size)
    c.drawRightString(PAGE_W - INNER_MARGIN_X, top_y, job_name)

    bot_y = BORDER_MARGIN + BOTTOM_MARGIN_Y

    c.setFont("Helvetica-Bold", ITEM_FONT_SIZE)
    c.drawString(INNER_MARGIN_X, bot_y, item_label)

    item_width = c.stringWidth(item_label, "Helvetica-Bold", ITEM_FONT_SIZE)
    max_desc_w = (PAGE_W - INNER_MARGIN_X) - (INNER_MARGIN_X + item_width + ITEM_DESC_GAP)
    max_desc_w = max(20, max_desc_w)

    c.setFont("Helvetica", DESC_FONT_SIZE)
    safe_desc = truncate_text(c, description, "Helvetica", DESC_FONT_SIZE, max_desc_w)
    c.drawRightString(PAGE_W - INNER_MARGIN_X, bot_y + 4, safe_desc)

    logo_reader = ImageReader(logo_path)
    lw, lh = logo_reader.getSize()

    logo_box_top = top_y - 18
    logo_box_bottom = bot_y + 18
    logo_box_h = max(1, logo_box_top - logo_box_bottom)
    logo_box_w = PAGE_W - 2 * INNER_MARGIN_X

    scale = min(logo_box_w / lw, logo_box_h / lh) * LOGO_SHRINK
    logo_w = lw * scale
    logo_h = lh * scale

    logo_x = (PAGE_W - logo_w) / 2
    logo_y = logo_box_bottom + (logo_box_h - logo_h) / 2

    c.drawImage(logo_path, logo_x, logo_y, width=logo_w, height=logo_h,
                preserveAspectRatio=True, mask="auto")


# ========================
# JOB + ITEMS PARSING
# ========================
def parse_job_header(text: str):
    """
    Works for both:
      Work Order: 119172 Diptyque SCP - Item 2 and 3
    and:
      Work Order: 118880
      Project Name: ...
    """
    job_number = " "
    job_name = "UNKNOWN JOB"

    m = re.search(r"Work Order:\s*([0-9]{5,6})(.*)$", text, re.IGNORECASE | re.MULTILINE)
    if m:
        job_number = f"#{m.group(1)}"
        tail = _norm(m.group(2))
        if tail:
            job_name = tail.strip(" -")
    pm = re.search(r"Project Name:\s*(.+)", text, re.IGNORECASE)
    if pm:
        job_name = _norm(pm.group(1))
    return job_number, job_name

def parse_items_any_format(text: str):
    """
    Parses item lines from:
      A) Work Order list table (your old format)
      B) Product Detail per-page header format:
         Item# Description Quantity Comments
         2.01 Open Tall w Divisions 1
    Returns: list of (item_label, desc, qty)
    """
    items = []
    seen = set()

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # Format B (Product Detail): "2.01 <desc> 1"  (qty is last integer)
    # This is very stable for the Product Detail report you uploaded.
    pd_re = re.compile(r"^(\d+\.\d+)\s+(.+?)\s+(\d+)\s*$")

    for ln in lines:
        # skip headers
        low = ln.lower()
        if low.startswith("item# description quantity"):
            continue
        if low.startswith("item#") and "description" in low and "quantity" in low:
            continue

        m = pd_re.match(ln)
        if m:
            item_no = m.group(1)
            raw_desc = m.group(2)
            qty = int(m.group(3))
            desc = clean_description_no_dims(raw_desc)

            key = (item_no, desc, qty)
            if key not in seen:
                seen.add(key)
                items.append((f"Item {item_no}", desc, qty))

    # If we found Product Detail items, that’s good enough
    if items:
        return items

    # Format A (fallback): "7.01 1 LH Dor 1 Drawer Base ..."
    # qty is second token, item is first token
    a_re = re.compile(r"^\s*(\d+\.\d+)\s+(\d+)\s+(.+?)\s*$")
    for ln in lines:
        m = a_re.match(ln)
        if not m:
            continue
        item_no = m.group(1)
        qty = int(m.group(2))
        raw_desc = m.group(3)
        desc = clean_description_no_dims(raw_desc)
        key = (item_no, desc, qty)
        if key not in seen:
            seen.add(key)
            items.append((f"Item {item_no}", desc, qty))

    return items


# ========================
# PARTS PARSING (SHELVES/TOE/BACK/DIV)
# ========================
# item anchors (Product Detail tables make item line easy to match)
ITEM_ANCHORS = [
    re.compile(r"^(\d+\.\d+)\s+.+?\s+\d+\s*$"),   # "2.01 Open Tall... 1"
    re.compile(r"\bItem\s*#?\s*[:\-]?\s*(\d+\.\d+)\b", re.IGNORECASE),
]

PART_ROW_RE = re.compile(r"^\s*(\d+)\s+(.+?)\s*$")  # qty + rest

def identify_part_type(name: str):
    t = (name or "").lower()

    # shelves
    if "shelf" in t and "clip" not in t:
        return "shelf"

    # toe kicks
    if "toe kick" in t or "toekick" in t or "toe-kick" in t:
        return "toe_kick"

    # backs
    if (" back" in f" {t}") and ("panel" in t or "back" in t or "backer" in t):
        return "back"

    # dividers: Microvellum uses "Division" a lot
    if "divider" in t or "division" in t or re.search(r"\bdiv\b", t):
        return "divider"

    return None

def parse_parts_from_product_detail(pdf_text: str):
    shelves_by_item = {}
    toe_by_item = {}
    back_by_item = {}
    div_by_item = {}

    current_item = None

    for raw_line in pdf_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # update current item if item anchor
        for rex in ITEM_ANCHORS:
            mm = rex.match(line)
            if mm:
                current_item = mm.group(1)
                break

        pm = PART_ROW_RE.match(line)
        if not pm or current_item is None:
            continue

        qty = int(pm.group(1))
        rest = pm.group(2)

        # remove dims/material/etc
        rest_clean = clean_description_no_dims(rest)
        part_type = identify_part_type(rest_clean)
        if not part_type:
            continue

        if part_type == "shelf":
            shelves_by_item[current_item] = shelves_by_item.get(current_item, 0) + qty
        elif part_type == "toe_kick":
            toe_by_item[current_item] = toe_by_item.get(current_item, 0) + qty
        elif part_type == "back":
            back_by_item[current_item] = back_by_item.get(current_item, 0) + qty
        elif part_type == "divider":
            div_by_item[current_item] = div_by_item.get(current_item, 0) + qty

    return shelves_by_item, toe_by_item, back_by_item, div_by_item

def shelf_set_labels(n: int):
    labels = []
    i = 1
    while i <= n:
        j = min(i + 1, n)
        if i == j:
            labels.append(f"SHELF {i}")
        else:
            labels.append(f"SHELVES {i}-{j}")
        i += 2
    return labels

def safe_item_sort_key(item_str: str):
    m = re.match(r"^\s*(\d+)\.(\d+)\s*$", (item_str or ""))
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return (10**9, 10**9, item_str or "")


# ========================
# GENERATE ONE PDF
# ========================
def generate_all_labels(job_number: str, job_name: str, items, pdf_text: str):
    logo_path = get_logo_path()
    output_path = pick_output_path(job_number, job_name)

    c = canvas.Canvas(output_path, pagesize=(PAGE_W, PAGE_H))

    # 1) item labels
    for item_label, description, qty in items:
        for _ in range(max(int(qty), 1)):
            draw_label(c, logo_path, job_number, job_name, item_label, description)
            c.showPage()

    # 2) parts labels (NO DIMS)
    shelves_by_item, toe_by_item, back_by_item, div_by_item = parse_parts_from_product_detail(pdf_text)

    # shelves: set of 2
    for item_no in sorted(shelves_by_item.keys(), key=safe_item_sort_key):
        n = int(shelves_by_item[item_no] or 0)
        for sdesc in shelf_set_labels(n):
            draw_label(c, logo_path, job_number, job_name, f"Item {item_no}", sdesc)
            c.showPage()

    # toe kicks
    for item_no in sorted(toe_by_item.keys(), key=safe_item_sort_key):
        n = int(toe_by_item[item_no] or 0)
        for _ in range(max(n, 0)):
            draw_label(c, logo_path, job_number, job_name, f"Item {item_no}", "TOE KICK")
            c.showPage()

    # backs
    for item_no in sorted(back_by_item.keys(), key=safe_item_sort_key):
        n = int(back_by_item[item_no] or 0)
        for _ in range(max(n, 0)):
            draw_label(c, logo_path, job_number, job_name, f"Item {item_no}", "BACK PANEL")
            c.showPage()

    # dividers
    for item_no in sorted(div_by_item.keys(), key=safe_item_sort_key):
        n = int(div_by_item[item_no] or 0)
        for _ in range(max(n, 0)):
            draw_label(c, logo_path, job_number, job_name, f"Item {item_no}", "DIVIDER")
            c.showPage()

    c.save()
    print("Saved to:", output_path)
    return output_path


# ========================
# MAIN
# ========================
def main():
    if len(sys.argv) < 2:
        print("Usage: python auto_labels_from_workorder.py <workorder.pdf | image>")
        return 2

    input_file = sys.argv[1].strip().strip('"')
    if not os.path.exists(input_file):
        print("File not found:", input_file)
        return 2

    if input_file.lower().endswith(".pdf"):
        pdf_text = extract_text_from_pdf(input_file)
        ocr_used = False
    else:
        pdf_text = extract_text_from_image(input_file)
        ocr_used = True

    job_number, job_name = parse_job_header(pdf_text)
    items = parse_items_any_format(pdf_text)

    print("OCR used:", ocr_used)
    print("Job:", job_number)
    print("Name:", job_name)
    print("Items found:", len(items))

    # IMPORTANT: even if items are missing, still generate parts labels (shelves etc.)
    if not items:
        print("Warning: No item header lines found; generating parts labels only (if any).")

    generate_all_labels(job_number, job_name, items, pdf_text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())