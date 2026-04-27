import os
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pdfplumber
import pytesseract
from PIL import Image, ImageEnhance, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


LABEL_PRESETS = {
    "4x6": {
        "page_size": landscape((4 * inch, 6 * inch)),
        "border_margin": 10,
        "inner_margin_x": 30,
        "top_margin_y": 26,
        "bottom_margin_y": 26,
        "job_font_size": 26,
        "item_font_size": 26,
        "desc_font_size": 18,
        "logo_shrink": 0.60,
        "item_desc_gap": 18,
        "line_width": 4,
        "job_name_max_offset": 120,
        "job_name_start_size": 24,
        "desc_y_offset": 4,
        "logo_top_gap": 18,
        "logo_bottom_gap": 18,
    },
    "2.5x6": {
        "page_size": landscape((2.5 * inch, 6 * inch)),
        "border_margin": 8,
        "inner_margin_x": 16,
        "top_margin_y": 18,
        "bottom_margin_y": 18,
        "job_font_size": 18,
        "item_font_size": 18,
        "desc_font_size": 12,
        "logo_shrink": 0.50,
        "item_desc_gap": 10,
        "line_width": 3,
        "job_name_max_offset": 70,
        "job_name_start_size": 12,
        "desc_y_offset": 2,
        "logo_top_gap": 10,
        "logo_bottom_gap": 14,
    },
}


@dataclass
class LabelRunConfig:
    output_dir: str
    logo_image_path: str | None = None
    logo_zip_path: str | None = None


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _signal_score(text: str) -> int:
    t = text or ""
    score = 0
    score += 30 if re.search(r"\bwork\s*order\b", t, re.IGNORECASE) else 0
    score += 20 if re.search(r"\bproject\s*name\b", t, re.IGNORECASE) else 0
    score += 12 * len(re.findall(r"\bitem\b", t, re.IGNORECASE))
    score += 8 * len(re.findall(rf"\b{ITEM_NO_RE}\b", t))
    score += min(20, int(len(re.findall(r"[A-Za-z]", t)) / 80))
    return score


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
        candidate = f"{base}_{i}{ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


def strip_repeated_job_number(job_number: str, job_name: str) -> str:
    name = (job_name or "").strip()
    digits = (job_number or "").replace("#", "").strip()

    if not digits or not name:
        return name

    patterns = [
        rf"^\s*#?\s*{re.escape(digits)}[\s\-_:]*",
        rf"^\s*job\s*#?\s*{re.escape(digits)}[\s\-_:]*",
    ]

    for pat in patterns:
        name = re.sub(pat, "", name, flags=re.IGNORECASE).strip()

    return name


def find_job_output_dir(job_number: str, output_root: str) -> str:
    os.makedirs(output_root, exist_ok=True)

    job_digits = (job_number or "").replace("#", "").strip()
    if not job_digits:
        fallback = os.path.join(output_root, "UNKNOWN_JOB")
        os.makedirs(fallback, exist_ok=True)
        return fallback

    for name in os.listdir(output_root):
        full = os.path.join(output_root, name)
        if os.path.isdir(full) and job_digits.lower() in name.lower():
            return full

    new_dir = os.path.join(output_root, job_digits)
    os.makedirs(new_dir, exist_ok=True)
    return new_dir


def build_output_base(job_number: str, job_name: str, output_root: str) -> str:
    job_digits = (job_number or "").replace("#", "").strip()
    job_name_clean = strip_repeated_job_number(job_number, job_name)

    safe_job = sanitize_filename(job_digits or "UNKNOWN")
    safe_name = sanitize_filename(job_name_clean or "UNKNOWN_JOB")

    job_dir = find_job_output_dir(job_number, output_root)
    return os.path.join(job_dir, f"{safe_job}_{safe_name}")


def pick_output_paths(job_number: str, job_name: str, output_root: str):
    base = build_output_base(job_number, job_name, output_root)
    path_4x6 = unique_path(f"{base}_Labels_4x6.pdf")
    path_25x6 = unique_path(f"{base}_Labels_2.5x6.pdf")
    return path_4x6, path_25x6


def extract_logo_from_zip(zip_path: str) -> str:
    extract_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__logo_extract__")
    os.makedirs(extract_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(extract_dir)

    for root, _, files in os.walk(extract_dir):
        for file_name in files:
            if file_name.lower().endswith((".png", ".jpg", ".jpeg", ".jfif", ".webp", ".bmp")):
                return os.path.join(root, file_name)

    raise RuntimeError("No logo image found inside logo zip.")


def get_logo_path(config: LabelRunConfig) -> str:
    if config.logo_image_path and os.path.exists(config.logo_image_path):
        return config.logo_image_path
    if config.logo_zip_path and os.path.exists(config.logo_zip_path):
        return extract_logo_from_zip(config.logo_zip_path)
    raise RuntimeError("Logo not found. Provide a logo image or zip.")


def _row_black_ratio(rgb_img: Image.Image, y: int, threshold: int = 22) -> float:
    row = rgb_img.crop((0, y, rgb_img.width, y + 1))
    pixels = list(row.getdata())
    black = 0
    for r, g, b in pixels:
        if r <= threshold and g <= threshold and b <= threshold:
            black += 1
    return black / max(1, len(pixels))


def prepare_logo_image(logo_path: str) -> str:
    """
    Trim top/bottom black bars (common when a screenshot is used as logo).
    Returns a path to a cleaned temporary PNG.
    """
    with Image.open(logo_path) as img:
        rgb = img.convert("RGB")
        w, h = rgb.size
        if h < 10 or w < 10:
            return logo_path

        max_trim = int(h * 0.45)
        black_row_cutoff = 0.93

        top = 0
        while top < min(max_trim, h - 1) and _row_black_ratio(rgb, top) >= black_row_cutoff:
            top += 1

        bottom = h - 1
        while bottom > max(0, h - 1 - max_trim) and _row_black_ratio(rgb, bottom) >= black_row_cutoff:
            bottom -= 1

        if bottom <= top:
            return logo_path

        cropped = rgb.crop((0, top, w, bottom + 1))

        cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "__logo_cache__")
        os.makedirs(cache_dir, exist_ok=True)
        out_path = os.path.join(cache_dir, "clean_logo.png")
        cropped.save(out_path, format="PNG")
        return out_path


def extract_text_from_pdf(path: str) -> str:
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            ocr_text = ""
            # Scanned/noisy PDFs often need OCR even when a weak text layer exists.
            if (not _norm(page_text)) or _signal_score(page_text) < 24:
                try:
                    rendered = page.to_image(resolution=300).original
                    ocr_text = _ocr_image_variants(rendered)
                except Exception:
                    ocr_text = ""

            if _signal_score(ocr_text) > _signal_score(page_text):
                page_text = ocr_text
            elif _norm(ocr_text) and _norm(ocr_text) != _norm(page_text):
                page_text = f"{page_text}\n{ocr_text}"
            text += page_text + "\n"
    return text


def extract_text_from_image(path: str) -> str:
    with Image.open(path) as image:
        return _ocr_image_variants(image)


def _ocr_image_variants(image: Image.Image) -> str:
    base = image.convert("RGB")
    scale = 2
    resized = base.resize((base.width * scale, base.height * scale), Image.Resampling.LANCZOS)
    gray = ImageOps.grayscale(resized)
    high_contrast = ImageEnhance.Contrast(gray).enhance(2.2)
    bw = high_contrast.point(lambda x: 0 if x < 165 else 255, mode="1")

    config = "--oem 1 --psm 6"
    candidates = [
        pytesseract.image_to_string(base, config=config),
        pytesseract.image_to_string(high_contrast, config=config),
        pytesseract.image_to_string(bw, config=config),
    ]
    return "\n".join(c for c in candidates if c and c.strip())


def _is_reasonable_description(text: str) -> bool:
    s = _norm(text)
    if len(s) < 2:
        return False
    alnum = sum(ch.isalnum() for ch in s)
    alpha = sum(ch.isalpha() for ch in s)
    if alnum == 0 or alpha < 2:
        return False
    if (alpha / alnum) < 0.32:
        return False
    # Reject obvious OCR gibberish dominated by punctuation fragments.
    if len(re.findall(r"[^A-Za-z0-9\s\-]", s)) > max(6, int(len(s) * 0.25)):
        return False
    return True


def _looks_like_metadata(text: str) -> bool:
    low = _norm(text).lower()
    if not low:
        return True
    meta_tokens = [
        "job_name",
        "job name",
        "job #",
        "job#",
        "date:",
        ".date",
        "page ",
        ".page",
        "page",
        "work order",
        "project name",
        "cut list",
        "print date",
    ]
    return any(token in low for token in meta_tokens)


NUM_TOKEN = r"(?:\d+(?:\.\d+)?)"
W_DIM_BLOCK_RE = re.compile(
    rf"\bw\s+(?:{NUM_TOKEN}\s+){{2,8}}(?:L|R|Left|Right|Both|None)?\b",
    re.IGNORECASE,
)
TRAIL_NUMS_RE = re.compile(
    rf"(?:\s+{NUM_TOKEN}){{1,10}}\s*(?:L|R|Left|Right|Both|None)?\s*$",
    re.IGNORECASE,
)
TRAIL_DIM_BLOCK_RE = re.compile(
    rf"(?:\s+{NUM_TOKEN}){{2,10}}\s*(?:L|R|LH|RH|Left|Right|Both|None)?\s*$",
    re.IGNORECASE,
)
FE_RE = re.compile(r"\bFE\s*[:=\-]?\s*[A-Za-z0-9]+\b", re.IGNORECASE)
ITEM_TOKEN_RE = re.compile(r"\bItem\s*\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?\b", re.IGNORECASE)


def _strip_trailing_quantity_markers(text: str) -> str:
    s = _norm(text)
    if not re.search(r"[A-Za-z]", s):
        return s

    patterns = [
        r"\s+[xX]\s*\d+\s*$",
        r"\s*[,;:\-]+\s*\d+\s*$",
        r"\s+\d+\s*$",
        r"\s+[.,]+\s*\d+\s*$",
        r"\s+\d+\s*(L|R|LH|RH|Left|Right|Both|None)\s*$",
    ]

    changed = True
    while changed:
        changed = False
        for pattern in patterns:
            updated = re.sub(pattern, "", s, flags=re.IGNORECASE)
            if updated != s:
                s = _norm(updated)
                changed = True
        s = re.sub(r"[\s,;:\-\.xX]+$", "", s)
        s = _norm(s)
    return s


def _is_dimension_token(token: str) -> bool:
    t = (token or "").strip().strip(",;:()[]{}")
    if not t:
        return False
    if re.fullmatch(r"[xX*/]", t):
        return True
    if re.fullmatch(r"\d+(?:\.\d+)?(?:\"|')?", t):
        return True
    if re.fullmatch(r"[WDHwdh]", t):
        return True
    if re.fullmatch(r"[WDHwdh]\d+(?:\.\d+)?", t):
        return True
    return False


def _strip_dimension_suffix(text: str) -> str:
    s = _norm(text)
    parts = s.split()
    if len(parts) < 3:
        return s

    for start in range(1, len(parts) - 1):
        suffix = parts[start:]
        if not _is_dimension_token(suffix[0]):
            continue
        dim_count = sum(1 for token in suffix if _is_dimension_token(token))
        non_dim_count = len(suffix) - dim_count

        # Remove a suffix only when it is overwhelmingly dimension-like.
        if dim_count >= 2 and non_dim_count == 0:
            return _norm(" ".join(parts[:start]))
        if dim_count >= 3 and non_dim_count <= 1:
            return _norm(" ".join(parts[:start]))

    return s


def clean_description_no_dims(raw: str) -> str:
    s = _norm(raw)
    s = ITEM_TOKEN_RE.sub("", s)
    s = _norm(s)
    # OCR often leaves a quantity marker before hand tags, e.g. "1 LH Door Upper".
    s = re.sub(r"^\s*\d+\s+(LH|RH|Left|Right)\b", r"\1", s, flags=re.IGNORECASE)
    s = _norm(s)
    s = FE_RE.sub("", s)
    s = _norm(s)
    s = W_DIM_BLOCK_RE.sub("", s)
    s = _norm(s)
    s = TRAIL_DIM_BLOCK_RE.sub("", s)
    s = _norm(s)
    s = _strip_dimension_suffix(s)
    s = _norm(s)
    s = TRAIL_NUMS_RE.sub("", s)
    s = _norm(s)
    s = _strip_trailing_quantity_markers(s)
    s = re.sub(r"\s+(L|R|Left|Right|Both|None)\s*$", "", s, flags=re.IGNORECASE)
    return _norm(s)


def truncate_text(cnv, text, font_name, font_size, max_width):
    if not text:
        return ""
    ellipsis = "..."
    if cnv.stringWidth(text, font_name, font_size) <= max_width:
        return text

    max_width = max(0, max_width - cnv.stringWidth(ellipsis, font_name, font_size))
    out = text
    while out and cnv.stringWidth(out, font_name, font_size) > max_width:
        out = out[:-1]
    return (out + ellipsis) if out else ellipsis


def choose_label_size(item_label: str, description: str) -> str:
    text = f"{item_label} {description}".lower()

    small_part_keywords = [
        "shelf",
        "shelves",
        "toe kick",
        "toekick",
        "toe-kick",
        "filler",
        "ceiling filler",
        "base filler",
        "upper filler",
        "light valance",
        "valance",
        "valence",
        "leg",
        "light rail",
        "countertop",
        "counter top",
        "ctop",
        "sub top",
        "subtop",
        "door hardware",
        "cork board",
        "mdf",
    ]

    for kw in small_part_keywords:
        if kw in text:
            return "2.5x6"

    return "4x6"


def draw_label(cnv, logo_path: str, job_number: str, job_name: str, item_label: str, description: str, label_size: str = "4x6"):
    cfg = LABEL_PRESETS[label_size]

    page_w, page_h = cfg["page_size"]
    border_margin = cfg["border_margin"]
    inner_margin_x = cfg["inner_margin_x"]
    top_margin_y = cfg["top_margin_y"]
    bottom_margin_y = cfg["bottom_margin_y"]
    job_font_size = cfg["job_font_size"]
    item_font_size = cfg["item_font_size"]
    desc_font_size = cfg["desc_font_size"]
    logo_shrink = cfg["logo_shrink"]
    item_desc_gap = cfg["item_desc_gap"]
    line_width = cfg["line_width"]
    job_name_max_offset = cfg["job_name_max_offset"]
    job_name_start_size = cfg["job_name_start_size"]
    desc_y_offset = cfg["desc_y_offset"]
    logo_top_gap = cfg["logo_top_gap"]
    logo_bottom_gap = cfg["logo_bottom_gap"]

    job_number = job_number or ""
    job_name = strip_repeated_job_number(job_number, job_name)
    item_label = item_label or ""
    description = description or ""

    cnv.setPageSize((page_w, page_h))
    cnv.setStrokeColor(colors.black)
    cnv.setLineWidth(line_width)
    cnv.rect(border_margin, border_margin, page_w - 2 * border_margin, page_h - 2 * border_margin)

    top_y = page_h - border_margin - top_margin_y

    cnv.setFont("Helvetica-Bold", job_font_size)
    cnv.drawString(inner_margin_x, top_y, job_number)

    name_size = job_name_start_size
    max_name_w = page_w - (2 * inner_margin_x) - job_name_max_offset
    while name_size > 8 and cnv.stringWidth(job_name, "Helvetica", name_size) > max_name_w:
        name_size -= 1

    cnv.setFont("Helvetica", name_size)
    cnv.drawRightString(page_w - inner_margin_x, top_y, job_name)

    bot_y = border_margin + bottom_margin_y

    cnv.setFont("Helvetica-Bold", item_font_size)
    cnv.drawString(inner_margin_x, bot_y, item_label)

    item_width = cnv.stringWidth(item_label, "Helvetica-Bold", item_font_size)
    max_desc_w = (page_w - inner_margin_x) - (inner_margin_x + item_width + item_desc_gap)
    max_desc_w = max(20, max_desc_w)

    cnv.setFont("Helvetica", desc_font_size)
    safe_desc = truncate_text(cnv, description, "Helvetica", desc_font_size, max_desc_w)
    cnv.drawRightString(page_w - inner_margin_x, bot_y + desc_y_offset, safe_desc)

    logo_reader = ImageReader(logo_path)
    lw, lh = logo_reader.getSize()

    logo_box_top = top_y - logo_top_gap
    logo_box_bottom = bot_y + logo_bottom_gap
    logo_box_h = max(1, logo_box_top - logo_box_bottom)
    logo_box_w = page_w - 2 * inner_margin_x

    scale = min(logo_box_w / lw, logo_box_h / lh) * logo_shrink
    logo_w = lw * scale
    logo_h = lh * scale

    logo_x = (page_w - logo_w) / 2
    logo_y = logo_box_bottom + (logo_box_h - logo_h) / 2

    cnv.drawImage(logo_path, logo_x, logo_y, width=logo_w, height=logo_h, preserveAspectRatio=True, mask="auto")


ITEM_NO_RE = r"\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?"


def parse_job_header(text: str):
    job_number = " "
    job_name = "UNKNOWN JOB"

    work_order_match = re.search(
        r"(?:Work\s*Order|Workorder)\s*[:#\-]?\s*([0-9]{4,8})(.*)$",
        text,
        re.IGNORECASE | re.MULTILINE,
    )
    if work_order_match:
        job_number = f"#{work_order_match.group(1)}"
        tail = _norm(work_order_match.group(2))
        if tail:
            job_name = tail.strip(" -")

    project_name_match = re.search(r"Project Name:\s*(.+)", text, re.IGNORECASE)
    if project_name_match:
        job_name = _norm(project_name_match.group(1))

    if not job_number.strip():
        fallback_wo = re.search(r"\b(?:WO|WO#|W/O|Order|Job)\s*[:#\-]?\s*(\d{4,8})\b", text, re.IGNORECASE)
        if fallback_wo:
            job_number = f"#{fallback_wo.group(1)}"
        else:
            any_digits = re.search(r"\b(\d{5,8})\b", text)
            if any_digits:
                job_number = f"#{any_digits.group(1)}"

    if job_name == "UNKNOWN JOB":
        for line in text.splitlines():
            ln = _norm(line)
            if re.search(r"\b(project|customer|client|name)\b", ln, re.IGNORECASE):
                cleaned = re.sub(r"^(Project\s*Name|Customer|Client|Name)\s*[:\-]?\s*", "", ln, flags=re.IGNORECASE)
                if _is_reasonable_description(cleaned):
                    job_name = cleaned
                    break

    return job_number, job_name


def parse_items_any_format(text: str):
    items = []
    seen = set()
    lines = [_norm(ln) for ln in text.splitlines() if _norm(ln)]

    # Normalize common OCR glitches before matching.
    normalized_lines = []
    for line in lines:
        fixed = re.sub(r"(?<=\d),(?=\d)", ".", line)  # 1,1 -> 1.1
        fixed = re.sub(r"\bltem\b", "Item", fixed, flags=re.IGNORECASE)  # ltem -> Item
        fixed = re.sub(r"\bItern\b", "Item", fixed, flags=re.IGNORECASE)
        normalized_lines.append(_norm(fixed))

    pd_re = re.compile(rf"^({ITEM_NO_RE})\s+(.+?)\s+(\d+)\s*$", re.IGNORECASE)
    item_desc_qty_re = re.compile(rf"^(?:Item\s*#?\s*)?({ITEM_NO_RE})\s+(.+?)\s+(\d+)\s*$", re.IGNORECASE)
    item_qty_desc_re = re.compile(rf"^(?:Item\s*#?\s*)?({ITEM_NO_RE})\s+(\d+)\s+(.+?)\s*$", re.IGNORECASE)
    item_desc_xqty_re = re.compile(rf"^(?:Item\s*#?\s*)?({ITEM_NO_RE})\s*[:\-]?\s+(.+?)\s*[xX]\s*(\d+)\s*$", re.IGNORECASE)
    qty_item_desc_re = re.compile(rf"^(\d+)\s+({ITEM_NO_RE})\s+(.+?)\s*$", re.IGNORECASE)

    for line in normalized_lines:
        low = line.lower()
        if low.startswith("item# description quantity"):
            continue
        if low.startswith("item#") and "description" in low and "quantity" in low:
            continue

        matchers = [pd_re, item_desc_qty_re, item_qty_desc_re, item_desc_xqty_re]
        found = None
        for rex in matchers:
            found = rex.match(line)
            if found:
                break

        if not found:
            swapped = qty_item_desc_re.match(line)
            if swapped:
                qty = int(swapped.group(1))
                item_no = swapped.group(2)
                raw_desc = swapped.group(3)
                desc = clean_description_no_dims(raw_desc)
                if _is_reasonable_description(desc) and 1 <= qty <= 200:
                    key = (item_no, desc, qty)
                    if key not in seen:
                        seen.add(key)
                        items.append((f"Item {item_no}", desc, qty))
            continue

        if not found:
            continue

        if rex is item_qty_desc_re:
            item_no = found.group(1)
            qty = int(found.group(2))
            raw_desc = found.group(3)
        else:
            item_no = found.group(1)
            raw_desc = found.group(2)
            qty = int(found.group(3))

        # Integer item numbers are noisy in OCR. Only trust them when line explicitly says "Item".
        if "." not in item_no and not re.search(r"\bitem\b", line, re.IGNORECASE):
            continue

        desc = clean_description_no_dims(raw_desc)
        if _looks_like_metadata(desc):
            continue
        if not _is_reasonable_description(desc):
            continue
        if qty < 1 or qty > 200:
            continue
        key = (item_no, desc, qty)
        if key not in seen:
            seen.add(key)
            items.append((f"Item {item_no}", desc, qty))

    if items:
        return items

    a_re = re.compile(rf"^\s*({ITEM_NO_RE})\s+(\d+)\s+(.+?)\s*$")
    for line in normalized_lines:
        m = a_re.match(line)
        if not m:
            continue

        item_no = m.group(1)
        qty = int(m.group(2))
        raw_desc = m.group(3)
        desc = clean_description_no_dims(raw_desc)
        if _looks_like_metadata(desc):
            continue
        if not _is_reasonable_description(desc):
            continue
        if qty < 1 or qty > 200:
            continue

        key = (item_no, desc, qty)
        if key not in seen:
            seen.add(key)
            items.append((f"Item {item_no}", desc, qty))

    if items:
        return items

    # Fallback for alternate cut list layouts with table-like spacing or extra columns.
    for line in normalized_lines:
        low = line.lower()
        if any(header in low for header in ["description", "quantity", "qty", "width", "height", "depth"]):
            continue

        cols = [c.strip() for c in re.split(r"\t+|\s{2,}", line) if c.strip()]
        candidates = cols if len(cols) >= 2 else [line]

        for candidate in candidates:
            match = re.search(rf"\b({ITEM_NO_RE})\b", candidate)
            if not match:
                continue

            item_no = match.group(1)
            before = candidate[:match.start()].strip()
            after = candidate[match.end():].strip()

            qty = None
            raw_desc = ""

            tail_qty = re.search(r"\b(\d{1,3})\s*$", after)
            if tail_qty:
                qty = int(tail_qty.group(1))
                raw_desc = after[:tail_qty.start()].strip()
            elif re.fullmatch(r"\d{1,3}", before):
                qty = int(before)
                raw_desc = after
            else:
                continue

            if not raw_desc:
                continue

            # Guard against the old false-positive "Item 1" collapse.
            if "." not in item_no and "-" not in item_no and not re.search(r"\bitem\b", candidate, re.IGNORECASE):
                continue

            desc = clean_description_no_dims(raw_desc)
            if _looks_like_metadata(desc):
                continue
            if not _is_reasonable_description(desc):
                continue
            if qty < 1 or qty > 200:
                continue

            key = (item_no, desc, qty)
            if key not in seen:
                seen.add(key)
                items.append((f"Item {item_no}", desc, qty))

    return items


def _split_page_sections(text: str):
    lines = [line.rstrip() for line in text.splitlines()]
    sections = []
    current = []

    for line in lines:
        if re.search(r"\bPAGE\b\s*[:#]?\s*\d+\b", line, re.IGNORECASE):
            if current:
                sections.append("\n".join(current))
                current = []
        current.append(line)

    if current:
        sections.append("\n".join(current))

    return [section for section in sections if _norm(section)]


def _extract_seed_from_second_row(text: str) -> int | None:
    lines = [_norm(line) for line in text.splitlines() if _norm(line)]
    if len(lines) < 2:
        return None
    second = lines[1]
    m = re.search(r"\b(\d{1,4})\b", second)
    if not m:
        return None
    return int(m.group(1))


def _extract_labeled_value(line: str) -> int | None:
    text = _norm(line)
    for pattern in [
        r"#\s*(\d{1,3})\b",
        r"\bno\.?\s*(\d{1,3})\b",
        r"\blabel\s*(\d{1,3})\b",
        r"\((\d{1,3})\)",
    ]:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return int(m.group(1))
    return None


def _parse_qty_from_line(text: str) -> int:
    m = re.match(r"^\s*(\d{1,3})\b", text)
    if m:
        return max(1, int(m.group(1)))
    m = re.search(r"\bqty\s*[:=]?\s*(\d{1,3})\b", text, re.IGNORECASE)
    if m:
        return max(1, int(m.group(1)))
    return 1


def _find_section_main_description(lines):
    keyword_scores = [
        ("upper cabinet", 9),
        ("lower cabinet", 9),
        ("base cabinet", 9),
        ("recycle cabinet", 9),
        ("full height panel", 9),
        ("tall cabinet", 8),
        ("wall cabinet", 8),
        ("cabinet", 8),
        ("panel", 7),
        ("upper", 6),
        ("base", 6),
        ("pantry", 6),
        ("drawer", 5),
        ("workstation", 5),
        ("kitchenette", 4),
    ]

    best_line = None
    best_score = -1
    for raw_line in lines:
        line = clean_description_no_dims(raw_line)
        low = line.lower()
        if _looks_like_metadata(line):
            continue
        if not _is_reasonable_description(line):
            continue
        if re.fullmatch(r"\([^)]*\)", line):
            continue
        score = 0
        for token, pts in keyword_scores:
            if token in low:
                score += pts
        if re.search(r"\b(front|side|left|right|back|rear|end)\b", low):
            score -= 4
        if any(token in low for token in ["filler", "valance", "light rail", "toe kick", "shelf", "sub top", "counter top", "countertop"]):
            score -= 6
        if len(line.split()) <= 1:
            score -= 2
        if score > best_score:
            best_score = score
            best_line = line

    if best_line:
        return best_line

    for raw_line in lines:
        line = clean_description_no_dims(raw_line)
        if not _looks_like_metadata(line) and _is_reasonable_description(line):
            return line
    return "UNKNOWN ITEM"


def _is_main_item_candidate(line: str) -> bool:
    cleaned = clean_description_no_dims(line)
    low = cleaned.lower()
    if not cleaned or _looks_like_metadata(cleaned):
        return False
    if not _is_reasonable_description(cleaned):
        return False
    if identify_part_type(cleaned):
        return False

    include_tokens = [
        "cabinet",
        "panel",
        "workstation",
        "work station",
        "pantry",
        "drawer base",
        "sink base",
        "microwave upper",
        "upper",
        "base",
        "recycle",
    ]
    return any(token in low for token in include_tokens)


def parse_sectioned_cut_list(text: str):
    sections = _split_page_sections(text)
    if len(sections) < 2:
        return [], None

    items = []
    part_details = {
        "shelf": {},
        "toe_kick": {},
        "filler": {},
        "valance": {},
        "light_rail": {},
        "sub_top": {},
        "back": {},
        "divider": {},
    }

    seed_major = _extract_seed_from_second_row(sections[0]) or _extract_seed_from_second_row(text) or 2
    item_index = 1
    used_item_nos = set()
    for section in sections:
        lines = [_norm(line) for line in section.splitlines() if _norm(line)]
        section_item_ids = []
        desc_seen = set()

        # First pass: discover all main item descriptions in this section.
        for line in lines:
            if not _is_main_item_candidate(line):
                continue
            desc = clean_description_no_dims(line)
            key = desc.lower()
            if key in desc_seen:
                continue
            desc_seen.add(key)
            label_value = _extract_labeled_value(line)
            synthetic_item_no = f"{seed_major}.{label_value:02d}" if label_value is not None else f"{seed_major}.{item_index:02d}"
            while synthetic_item_no in used_item_nos:
                item_index += 1
                synthetic_item_no = f"{seed_major}.{item_index:02d}"
            used_item_nos.add(synthetic_item_no)
            items.append((f"Item {synthetic_item_no}", desc, 1))
            section_item_ids.append(synthetic_item_no)
            item_index += 1

        if not section_item_ids:
            main_desc = _find_section_main_description(lines)
            if main_desc != "UNKNOWN ITEM":
                synthetic_item_no = f"{seed_major}.{item_index:02d}"
                while synthetic_item_no in used_item_nos:
                    item_index += 1
                    synthetic_item_no = f"{seed_major}.{item_index:02d}"
                used_item_nos.add(synthetic_item_no)
                items.append((f"Item {synthetic_item_no}", main_desc, 1))
                section_item_ids.append(synthetic_item_no)
                item_index += 1

        if not section_item_ids:
            continue

        # Second pass: attach small parts to the nearest item in this section.
        active_item = section_item_ids[0]
        for line in lines:
            if _is_main_item_candidate(line):
                candidate_desc = clean_description_no_dims(line).lower()
                for item_id, item_desc, _ in items:
                    if item_id in section_item_ids and item_desc.lower() == candidate_desc:
                        active_item = item_id.split("Item ", 1)[1] if item_id.startswith("Item ") else item_id
                        break

            cleaned = clean_description_no_dims(line)
            if _looks_like_metadata(cleaned) or not _is_reasonable_description(cleaned):
                continue
            part_type = identify_part_type(cleaned)
            if not part_type:
                continue
            qty = _parse_qty_from_line(line)
            part_details[part_type].setdefault(active_item, []).append((cleaned, qty))

    if len(items) < 2:
        return [], None

    return items, part_details


ITEM_ANCHORS = [
    re.compile(r"^(\d+\.\d+)\s+.+?\s+\d+\s*$"),
    re.compile(rf"\bItem\s*#?\s*[:\-]?\s*({ITEM_NO_RE})\b", re.IGNORECASE),
]
PART_ROW_RE = re.compile(r"^\s*(\d+)\s+(.+?)\s*$")


def identify_part_type(name: str):
    text = (name or "").lower()

    if "shelf" in text and "clip" not in text:
        return "shelf"
    if "toe kick" in text or "toekick" in text or "toe-kick" in text:
        return "toe_kick"
    if "filler" in text:
        return "filler"
    if "valance" in text or "valence" in text:
        return "valance"
    if "light rail" in text:
        return "light_rail"
    if "sub top" in text or "subtop" in text or "counter top" in text or "countertop" in text:
        return "sub_top"
    # Keep this strict to avoid false positives from OCR noise in sectioned cut lists.
    if re.search(r"\b(back\s*panel|panel\s*back|backer\s*panel|cabinet\s*back|back\s*board)\b", text):
        return "back"
    if "divider" in text or "division" in text or re.search(r"\bdiv\b", text):
        return "divider"

    return None


def parse_part_details_from_product_detail(pdf_text: str):
    details = {
        "shelf": {},
        "toe_kick": {},
        "filler": {},
        "valance": {},
        "light_rail": {},
        "sub_top": {},
        "back": {},
        "divider": {},
    }

    current_item = None

    for raw_line in pdf_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

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

        rest_clean = clean_description_no_dims(rest)
        part_type = identify_part_type(rest_clean)
        if not part_type:
            continue

        bucket = details[part_type].setdefault(current_item, [])
        bucket.append((rest_clean, qty))

    return details


def parse_parts_from_product_detail(pdf_text: str):
    shelves_by_item = {}
    toe_by_item = {}
    filler_by_item = {}
    valance_by_item = {}
    light_rail_by_item = {}
    sub_top_by_item = {}
    back_by_item = {}
    div_by_item = {}

    part_details = parse_part_details_from_product_detail(pdf_text)

    for item_no, rows in part_details["shelf"].items():
        shelves_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["toe_kick"].items():
        toe_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["filler"].items():
        filler_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["valance"].items():
        valance_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["light_rail"].items():
        light_rail_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["sub_top"].items():
        sub_top_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["back"].items():
        back_by_item[item_no] = sum(qty for _, qty in rows)
    for item_no, rows in part_details["divider"].items():
        div_by_item[item_no] = sum(qty for _, qty in rows)

    return (
        shelves_by_item,
        toe_by_item,
        filler_by_item,
        valance_by_item,
        light_rail_by_item,
        sub_top_by_item,
        back_by_item,
        div_by_item,
    )


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
    m = re.match(r"^\s*(\d+)(?:\.(\d+))?(?:-(\d+)(?:\.(\d+))?)?\s*$", (item_str or ""))
    if m:
        major = int(m.group(1))
        minor = int(m.group(2) or 0)
        range_major = int(m.group(3) or 0)
        range_minor = int(m.group(4) or 0)
        return (major, minor, range_major, range_minor)
    return (10**9, 10**9, item_str or "")


def extract_item_number(item_label: str) -> str | None:
    m = re.search(rf"\b({ITEM_NO_RE})\b", item_label or "")
    return m.group(1) if m else None


def build_sequential_item_map(items, part_details):
    source_item_numbers = set()

    for item_label, _, _ in items:
        item_no = extract_item_number(item_label)
        if item_no:
            source_item_numbers.add(item_no)

    for bucket in part_details.values():
        source_item_numbers.update(bucket.keys())

    ordered = sorted(source_item_numbers, key=safe_item_sort_key)
    mapping = {}
    counter = 201
    for item_no in ordered:
        mapping[item_no] = f"{counter // 100}.{counter % 100:02d}"
        counter += 1
    return mapping


def _find_part_orientation(text: str) -> str | None:
    low = (text or "").lower()
    for token in ["front", "side", "left", "right", "back", "rear", "end"]:
        if re.search(rf"\b{token}\b", low):
            return "back" if token == "rear" else token
    return None


def build_small_part_label(part_type: str, raw_desc: str, index: int) -> str:
    desc = clean_description_no_dims(raw_desc)
    desc = re.sub(r"^\s*\d+\s+", "", desc)
    orientation = _find_part_orientation(desc)
    low = desc.lower()

    if part_type == "filler":
        parts = [f"filler #{index}"]
        if orientation:
            parts.append(orientation)
        return " ".join(parts)

    if part_type == "valance":
        base = "light valance" if "light" in low else "valance"
        parts = [f"{base} #{index}"]
        if orientation:
            parts.append(orientation)
        return " ".join(parts)

    if part_type == "light_rail":
        parts = [f"light rail #{index}"]
        if orientation:
            parts.append(orientation)
        return " ".join(parts)

    if part_type == "sub_top":
        return "counter top" if "counter" in low else "sub top"

    if part_type == "toe_kick":
        return "toe kick"

    return desc


def build_small_part_group_key(part_type: str, raw_desc: str) -> str:
    desc = clean_description_no_dims(raw_desc).lower()
    desc = re.sub(r"^\s*\d+\s+", "", desc)
    desc = re.sub(r"\b(front|side|left|right|back|rear|end)\b", "", desc)
    desc = re.sub(r"\s+", " ", desc).strip()
    return f"{part_type}:{desc}"


def normalize_small_part_description(text: str) -> str:
    desc = clean_description_no_dims(text)
    desc = re.sub(r"^\s*\d+\s+", "", desc)
    return _norm(desc)


def collapse_part_rows(rows, part_type: str):
    grouped = {}
    order = []
    for raw_desc, qty in rows:
        key = build_small_part_group_key(part_type, raw_desc)
        if key not in grouped:
            grouped[key] = [raw_desc, qty]
            order.append(key)
        else:
            grouped[key][1] = max(grouped[key][1], qty)
    return [(grouped[key][0], grouped[key][1]) for key in order]


def generate_all_labels(
    job_number: str,
    job_name: str,
    items,
    pdf_text: str,
    config: LabelRunConfig,
    part_details_override=None,
    use_sequential_item_numbers: bool = False,
):
    logo_path = prepare_logo_image(get_logo_path(config))
    output_4x6, output_25x6 = pick_output_paths(job_number, job_name, config.output_dir)
    part_details = part_details_override or parse_part_details_from_product_detail(pdf_text)
    if use_sequential_item_numbers:
        item_map = build_sequential_item_map(items, part_details)
    else:
        item_map = {}
        for item_label, _, _ in items:
            source_item_no = extract_item_number(item_label)
            if source_item_no:
                item_map[source_item_no] = source_item_no
        for bucket in part_details.values():
            for item_no in bucket.keys():
                item_map.setdefault(item_no, item_no)

    c4 = canvas.Canvas(output_4x6, pagesize=LABEL_PRESETS["4x6"]["page_size"])
    c25 = canvas.Canvas(output_25x6, pagesize=LABEL_PRESETS["2.5x6"]["page_size"])

    count_4x6 = 0
    count_25x6 = 0

    def route_label(item_label: str, description: str, qty: int = 1):
        nonlocal count_4x6, count_25x6

        safe_description = clean_description_no_dims(description)
        label_size = choose_label_size(item_label, safe_description)
        target = c25 if label_size == "2.5x6" else c4

        for _ in range(max(int(qty), 1)):
            draw_label(
                target,
                logo_path,
                job_number,
                job_name,
                item_label,
                safe_description,
                label_size=label_size,
            )
            target.showPage()
            if label_size == "2.5x6":
                count_25x6 += 1
            else:
                count_4x6 += 1

    for item_label, description, qty in items:
        source_item_no = extract_item_number(item_label)
        mapped_item = f"Item {item_map.get(source_item_no, source_item_no or '1.01')}"
        normalized_desc = normalize_small_part_description(description)
        part_type = identify_part_type(normalized_desc)
        effective_qty = qty
        effective_desc = description

        # Small parts embedded in the main item stream should not explode by dimension-derived qty.
        if part_type in {"filler", "valance", "light_rail", "sub_top"}:
            effective_qty = 1
            effective_desc = normalized_desc

        route_label(mapped_item, effective_desc, effective_qty)

    if part_details_override is None:
        (
            shelves_by_item,
            toe_by_item,
            filler_by_item,
            valance_by_item,
            light_rail_by_item,
            sub_top_by_item,
            back_by_item,
            div_by_item,
        ) = parse_parts_from_product_detail(pdf_text)
    else:
        shelves_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["shelf"].items()}
        toe_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["toe_kick"].items()}
        filler_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["filler"].items()}
        valance_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["valance"].items()}
        light_rail_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["light_rail"].items()}
        sub_top_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["sub_top"].items()}
        back_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["back"].items()}
        div_by_item = {k: sum(qty for _, qty in v) for k, v in part_details["divider"].items()}

    for item_no in sorted(shelves_by_item.keys(), key=safe_item_sort_key):
        n = int(shelves_by_item[item_no] or 0)
        for sdesc in shelf_set_labels(n):
            route_label(f"Item {item_map.get(item_no, item_no)}", sdesc, 1)

    for item_no in sorted(toe_by_item.keys(), key=safe_item_sort_key):
        route_label(f"Item {item_map.get(item_no, item_no)}", "toe kick", int(toe_by_item[item_no] or 0))

    for item_no in sorted(part_details["filler"].keys(), key=safe_item_sort_key):
        mapped_item = f"Item {item_map.get(item_no, item_no)}"
        filler_index_by_group = {}
        next_index = 1
        for raw_desc, qty in collapse_part_rows(part_details["filler"][item_no], "filler"):
            group_key = build_small_part_group_key("filler", raw_desc)
            if group_key not in filler_index_by_group:
                filler_index_by_group[group_key] = next_index
                next_index += 1
            idx = filler_index_by_group[group_key]
            route_label(mapped_item, build_small_part_label("filler", raw_desc, idx), max(qty, 1))

    for item_no in sorted(part_details["valance"].keys(), key=safe_item_sort_key):
        mapped_item = f"Item {item_map.get(item_no, item_no)}"
        valance_index_by_group = {}
        next_index = 1
        for raw_desc, qty in collapse_part_rows(part_details["valance"][item_no], "valance"):
            group_key = build_small_part_group_key("valance", raw_desc)
            if group_key not in valance_index_by_group:
                valance_index_by_group[group_key] = next_index
                next_index += 1
            idx = valance_index_by_group[group_key]
            route_label(mapped_item, build_small_part_label("valance", raw_desc, idx), max(qty, 1))

    for item_no in sorted(part_details["light_rail"].keys(), key=safe_item_sort_key):
        mapped_item = f"Item {item_map.get(item_no, item_no)}"
        rail_index_by_group = {}
        next_index = 1
        for raw_desc, qty in collapse_part_rows(part_details["light_rail"][item_no], "light_rail"):
            group_key = build_small_part_group_key("light_rail", raw_desc)
            if group_key not in rail_index_by_group:
                rail_index_by_group[group_key] = next_index
                next_index += 1
            idx = rail_index_by_group[group_key]
            route_label(mapped_item, build_small_part_label("light_rail", raw_desc, idx), max(qty, 1))

    for item_no in sorted(part_details["sub_top"].keys(), key=safe_item_sort_key):
        mapped_item = f"Item {item_map.get(item_no, item_no)}"
        for idx, (raw_desc, qty) in enumerate(collapse_part_rows(part_details["sub_top"][item_no], "sub_top"), start=1):
            route_label(mapped_item, build_small_part_label("sub_top", raw_desc, idx), max(qty, 1))

    for item_no in sorted(back_by_item.keys(), key=safe_item_sort_key):
        route_label(f"Item {item_map.get(item_no, item_no)}", "BACK PANEL", int(back_by_item[item_no] or 0))

    for item_no in sorted(div_by_item.keys(), key=safe_item_sort_key):
        route_label(f"Item {item_map.get(item_no, item_no)}", "DIVIDER", int(div_by_item[item_no] or 0))

    saved_paths = []

    if count_4x6 > 0:
        c4.save()
        saved_paths.append(output_4x6)
    else:
        try:
            os.remove(output_4x6)
        except FileNotFoundError:
            pass

    if count_25x6 > 0:
        c25.save()
        saved_paths.append(output_25x6)
    else:
        try:
            os.remove(output_25x6)
        except FileNotFoundError:
            pass

    return saved_paths


def generate_manual_labels(
    config: LabelRunConfig,
    job_number: str,
    job_name: str,
    item_label: str,
    description: str,
    quantity: int = 1,
    label_size: str = "auto",
):
    logo_path = prepare_logo_image(get_logo_path(config))
    output_4x6, output_25x6 = pick_output_paths(job_number, job_name, config.output_dir)

    c4 = canvas.Canvas(output_4x6, pagesize=LABEL_PRESETS["4x6"]["page_size"])
    c25 = canvas.Canvas(output_25x6, pagesize=LABEL_PRESETS["2.5x6"]["page_size"])

    chosen_size = label_size if label_size in {"4x6", "2.5x6"} else choose_label_size(item_label, description)
    safe_description = clean_description_no_dims(description)
    chosen_size = label_size if label_size in {"4x6", "2.5x6"} else choose_label_size(item_label, safe_description)
    qty = max(int(quantity or 1), 1)

    target = c25 if chosen_size == "2.5x6" else c4
    for _ in range(qty):
        draw_label(
            target,
            logo_path,
            job_number or " ",
            job_name or "UNKNOWN JOB",
            item_label or "Item 1",
            safe_description or "",
            label_size=chosen_size,
        )
        target.showPage()

    saved_paths = []
    if chosen_size == "4x6":
        c4.save()
        saved_paths.append(output_4x6)
        try:
            os.remove(output_25x6)
        except FileNotFoundError:
            pass
    else:
        c25.save()
        saved_paths.append(output_25x6)
        try:
            os.remove(output_4x6)
        except FileNotFoundError:
            pass

    return saved_paths


def run_label_generation(
    input_file: str,
    config: LabelRunConfig,
    job_number_override: Optional[str] = None,
    job_name_override: Optional[str] = None,
):
    path = Path(input_file)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    suffix = path.suffix.lower()
    if suffix == ".pdf":
        pdf_text = extract_text_from_pdf(str(path))
        ocr_used = False
    else:
        pdf_text = extract_text_from_image(str(path))
        ocr_used = True

    job_number, job_name = parse_job_header(pdf_text)
    if job_number_override and _norm(job_number_override):
        raw = _norm(job_number_override).replace(" ", "")
        job_number = raw if raw.startswith("#") else f"#{raw}"
    if job_name_override and _norm(job_name_override):
        job_name = _norm(job_name_override)
    items = parse_items_any_format(pdf_text)
    part_details_override = None
    section_items, section_parts = parse_sectioned_cut_list(pdf_text)
    if len(items) <= 1 and section_items:
        items = section_items
        part_details_override = section_parts

    parsed_parts = parse_parts_from_product_detail(pdf_text) if part_details_override is None else (
        {k: sum(qty for _, qty in v) for k, v in part_details_override["shelf"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["toe_kick"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["filler"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["valance"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["light_rail"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["sub_top"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["back"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["divider"].items()},
    )
    parsed_parts_total = sum(sum(bucket.values()) for bucket in parsed_parts)

    outputs = generate_all_labels(
        job_number,
        job_name,
        items,
        pdf_text,
        config,
        part_details_override=part_details_override,
        use_sequential_item_numbers=False,
    )

    return {
        "ocr_used": ocr_used,
        "job_number": job_number,
        "job_name": job_name,
        "items_found": len(items),
        "parsed_parts_total": parsed_parts_total,
        "text_chars": len(pdf_text or ""),
        "output_files": outputs,
    }


def _parse_input_to_components(
    input_file: str,
    job_number_override: Optional[str] = None,
    job_name_override: Optional[str] = None,
):
    path = Path(input_file)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {input_file}")

    suffix = path.suffix.lower()
    if suffix == ".pdf":
        pdf_text = extract_text_from_pdf(str(path))
        ocr_used = False
    else:
        pdf_text = extract_text_from_image(str(path))
        ocr_used = True

    job_number, job_name = parse_job_header(pdf_text)
    if job_number_override and _norm(job_number_override):
        raw = _norm(job_number_override).replace(" ", "")
        job_number = raw if raw.startswith("#") else f"#{raw}"
    if job_name_override and _norm(job_name_override):
        job_name = _norm(job_name_override)

    items = parse_items_any_format(pdf_text)
    part_details_override = None
    section_items, section_parts = parse_sectioned_cut_list(pdf_text)
    if len(items) <= 1 and section_items:
        items = section_items
        part_details_override = section_parts

    parsed_parts = parse_parts_from_product_detail(pdf_text) if part_details_override is None else (
        {k: sum(qty for _, qty in v) for k, v in part_details_override["shelf"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["toe_kick"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["filler"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["valance"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["light_rail"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["sub_top"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["back"].items()},
        {k: sum(qty for _, qty in v) for k, v in part_details_override["divider"].items()},
    )
    parsed_parts_total = sum(sum(bucket.values()) for bucket in parsed_parts)

    return {
        "ocr_used": ocr_used,
        "job_number": job_number,
        "job_name": job_name,
        "items": items,
        "part_details_override": part_details_override,
        "parsed_parts_total": parsed_parts_total,
        "text_chars": len(pdf_text or ""),
        "pdf_text": pdf_text,
    }


def _empty_part_details():
    return {
        "shelf": {},
        "toe_kick": {},
        "filler": {},
        "valance": {},
        "light_rail": {},
        "sub_top": {},
        "back": {},
        "divider": {},
    }


def _merge_part_details(merged, incoming):
    for part_type, item_map in incoming.items():
        target_item_map = merged.setdefault(part_type, {})
        for item_no, rows in item_map.items():
            row_map = target_item_map.setdefault(item_no, {})
            for raw_desc, qty in rows:
                cleaned = clean_description_no_dims(raw_desc)
                key = _norm(cleaned).lower()
                if not key:
                    continue
                row_map[key] = max(row_map.get(key, 0), int(qty or 0))


def _finalize_merged_part_details(merged):
    finalized = _empty_part_details()
    for part_type, item_map in merged.items():
        for item_no, row_map in item_map.items():
            rows = [(desc, qty) for desc, qty in row_map.items() if qty > 0]
            if rows:
                finalized[part_type][item_no] = rows
    return finalized


def run_multi_label_generation(
    input_files: list[str],
    config: LabelRunConfig,
    job_number_override: Optional[str] = None,
    job_name_override: Optional[str] = None,
):
    if not input_files:
        raise ValueError("No input files provided.")

    combined_items_map = {}
    merged_part_details = _empty_part_details()
    combined_text_parts = []
    ocr_used_any = False
    total_text_chars = 0
    total_parsed_parts = 0
    resolved_job_number = None
    resolved_job_name = None

    for input_file in input_files:
        parsed = _parse_input_to_components(
            input_file,
            job_number_override=job_number_override,
            job_name_override=job_name_override,
        )
        ocr_used_any = ocr_used_any or parsed["ocr_used"]
        total_text_chars += int(parsed["text_chars"] or 0)
        total_parsed_parts += int(parsed["parsed_parts_total"] or 0)
        combined_text_parts.append(parsed["pdf_text"] or "")

        if not resolved_job_number and _norm(parsed["job_number"]):
            resolved_job_number = parsed["job_number"]
        if (not resolved_job_name or resolved_job_name == "UNKNOWN JOB") and _norm(parsed["job_name"]):
            resolved_job_name = parsed["job_name"]

        for item_label, description, qty in parsed["items"]:
            clean_label = _norm(item_label)
            clean_desc = _norm(clean_description_no_dims(description))
            if not clean_label:
                continue
            key = (clean_label.lower(), clean_desc.lower())
            combined_items_map[key] = (clean_label, clean_desc, max(combined_items_map.get(key, ("", "", 0))[2], int(qty or 0)))

        incoming_part_details = parsed["part_details_override"] or parse_part_details_from_product_detail(parsed["pdf_text"])
        _merge_part_details(merged_part_details, incoming_part_details)

    merged_items = list(combined_items_map.values())
    merged_items.sort(key=lambda row: safe_item_sort_key(extract_item_number(row[0]) or ""))
    finalized_part_details = _finalize_merged_part_details(merged_part_details)

    job_number = resolved_job_number or " "
    job_name = resolved_job_name or "UNKNOWN JOB"

    outputs = generate_all_labels(
        job_number,
        job_name,
        merged_items,
        "\n\n".join(combined_text_parts),
        config,
        part_details_override=finalized_part_details,
        use_sequential_item_numbers=False,
    )

    return {
        "ocr_used": ocr_used_any,
        "job_number": job_number,
        "job_name": job_name,
        "items_found": len(merged_items),
        "parsed_parts_total": total_parsed_parts,
        "text_chars": total_text_chars,
        "output_files": outputs,
        "files_processed": len(input_files),
    }
