import os, uuid, time, json, re, sys
from pathlib import Path
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt, Cm, Inches, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4
from dotenv import load_dotenv
from google import genai
from google.genai import types as genai_types

load_dotenv(override=True)

# Force UTF-8 output so Greek/math symbols in OCR results don't crash the console
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass  # Python < 3.7 or non-TextIOWrapper stdout

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
from pydantic import BaseModel, Field
from typing import List, Optional

class DocElement(BaseModel):
    type: str = Field(
        ...,
        description="The type of the element. Must be one of: 'text', 'blank_line', 'table', 'drawing'."
    )
    text: Optional[str] = Field(
        None,
        description="The transcribed text content of the line or paragraph."
    )
    tag: Optional[str] = Field(
        None,
        description="The formatting tag for text elements: 'HEADING' (major title/heading), 'SUBHEAD' (sub-heading/section label), 'BODY' (normal body paragraph), 'BULLET' (bullet point or numbered list item), 'CENTER' (visually centered line), 'UNDERLN' (underlined text)."
    )
    bold: Optional[bool] = Field(
        None,
        description="True if the text is bold or written significantly darker/thicker than normal."
    )
    italic: Optional[bool] = Field(
        None,
        description="True if the text is italicized or slanted."
    )
    underline: Optional[bool] = Field(
        None,
        description="True if the text is underlined."
    )
    alignment: Optional[str] = Field(
        None,
        description="Text alignment: 'left', 'center', 'right', 'justify'."
    )
    left_indent_cm: Optional[float] = Field(
        None,
        description="Estimated left indentation of the text element in centimeters (usually 0.0, or 0.5 for bullet points/indented blocks)."
    )
    table_data: Optional[List[List[str]]] = Field(
        None,
        description="A list of rows, where each row is a list of strings representing the cells of the table."
    )
    bbox: Optional[List[float]] = Field(
        None,
        description="Normalized bounding box [x1, y1, x2, y2] (from 0.0 to 1.0) of the drawing or diagram in the image."
    )
    description: Optional[str] = Field(
        None,
        description="A one-sentence description of the drawing, diagram, sketch, or graph."
    )
    is_simple_arrow: Optional[bool] = Field(
        None,
        description="True if the entire drawing is just a single plain hand-drawn arrow (no labels, text, or boxes)."
    )
    arrow_direction: Optional[str] = Field(
        None,
        description="If is_simple_arrow is True, the direction the arrow points: 'right', 'left', 'up', 'down', 'up-right', 'up-left', 'down-right', 'down-left'."
    )
    is_simple_bracket: Optional[bool] = Field(
        None,
        description="True if the entire drawing is just a single grouping bracket or curly brace."
    )
    bracket_style: Optional[str] = Field(
        None,
        description="If is_simple_bracket is True, the style of bracket: 'curly', 'square', 'plain'."
    )
    bracket_side: Optional[str] = Field(
        None,
        description="If is_simple_bracket is True, which side the bracket appears relative to the text: 'left', 'right'."
    )

class DocumentLayout(BaseModel):
    page_margin_cm: float = Field(
        default=2.54,
        description="Estimated page margin in centimeters (usually between 1.5 and 3.0, default 2.54)."
    )
    line_spacing: float = Field(
        default=1.15,
        description="Line spacing: 1.0 (tight), 1.15 (normal), 1.5 (airy), 2.0 (double-spaced)."
    )
    elements: List[DocElement] = Field(
        ...,
        description="The ordered list of all document elements (text, tables, drawings, blank lines) from top to bottom of the page in reading order."
    )

client = genai.Client(
    api_key=GEMINI_API_KEY,
    http_options=genai_types.HttpOptions(timeout=120_000),
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

ALIGN_MAP = {
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# ── Pages ──
@app.get("/", response_class=HTMLResponse)
def landing(request: Request):
    return templates.TemplateResponse("landing.html", {"request": request})

@app.get("/convert", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/login", response_class=HTMLResponse)
def login(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})

@app.get("/history", response_class=HTMLResponse)
def history(request: Request):
    return templates.TemplateResponse("history.html", {"request": request})

@app.get("/about", response_class=HTMLResponse)
def about(request: Request):
    return templates.TemplateResponse("about.html", {"request": request})


# ── PDF → Images ──
def pdf_to_images(path):
    imgs = convert_from_path(path)
    paths = []
    for i, img in enumerate(imgs):
        p = f"{UPLOAD_DIR}/{uuid.uuid4()}_{i}.png"
        img.save(p)
        paths.append(p)
    return paths


# ═══════════════════════════════════════════════════════════════════════
#  COMBINED OCR PROMPT — Consolidated OCR & Layout structured detection
# ═══════════════════════════════════════════════════════════════════════
COMBINED_OCR_PROMPT = """\
You are an expert document analyzer and handwriting OCR engine. Your task is to perform high-fidelity text transcription and document layout analysis on the provided image, returning the results in a structured format.

Analyse the document from TOP to BOTTOM, LEFT to RIGHT. Generate a list of elements in their exact reading order.

━━━ TRANSCRIPTION RULES ━━━
1. Transcribe text with 100% fidelity. Do not correct spelling, grammar, or abbreviations.
2. If a word is completely illegible, transcribe it as [?].
3. Preserve math symbols, subscripts (x² -> x^2), fractions (3/4), and equations.
4. For text elements, determine the tag:
   - HEADING: Major title (significantly larger/bolder)
   - SUBHEAD: Sub-heading or section label (bold, slightly larger)
   - BODY: Normal paragraph or text line
   - BULLET: List item or bullet point
   - CENTER: Centered text (e.g. date, page title)
   - UNDERLN: Text that has an underline drawn directly under it.
5. For text elements, detect bold, italic, and underline formatting, and text alignment ('left', 'center', 'right', 'justify').
6. Estimate the left indentation of text in centimeters (e.g., 0.5 cm for bullet points or indented paragraphs).

━━━ NATIVE TABLES ━━━
Identify any tables in the image. For each table, return it as a 'table' element with `table_data` representing a 2D grid of cell texts. Extract all cell text accurately.

━━━ DRAWINGS, DIAGRAMS & SHAPES ━━━
Identify all drawings (diagrams, flowcharts, graphs, arrows, brackets, sketches).
For each drawing, return a 'drawing' element with a normalized bounding box [x1, y1, x2, y2] where x1,y1 is top-left and x2,y2 is bottom-right as fractions of image width and height (values from 0.0 to 1.0).
- Simple Arrow: If the drawing is only a simple arrow (e.g. pointing from one item to another), set is_simple_arrow=True and specify arrow_direction ('right', 'left', 'up', 'down', 'up-right', 'up-left', 'down-right', 'down-left').
- Simple Bracket: If the drawing is only a grouping bracket or brace (curly '{', square '[', or plain L-shaped), set is_simple_bracket=True, specify bracket_style ('curly', 'square', 'plain') and bracket_side ('left' or 'right' of the text).
- General Drawings: For charts, flowcharts, or sketches, set is_simple_arrow and is_simple_bracket to False and provide a brief description.

━━━ PAGE SETTINGS ━━━
Estimate the overall page margins in cm (page_margin_cm) and the line spacing (line_spacing: 1.0 for tight, 1.15 for normal, 1.5 for airy, 2.0 for double-spaced).
"""

PLAIN_OCR_PROMPT = (
    "Extract all the text from this handwriting image exactly as written. "
    "Read every line top to bottom. Output each line on its own line. "
    "Preserve capitalisation, punctuation, numbers and symbols exactly. "
    "Skip any diagrams or drawings silently. "
    "No descriptions, no extra formatting, no commentary."
)


# ─────────────────────────────────────────────────
#  Image preprocessing — enhance for OCR accuracy
# ─────────────────────────────────────────────────
def _preprocess_for_ocr(img: Image.Image) -> Image.Image:
    """
    Standardize the image for Gemini multimodal analysis:
    - Ensure it is in RGB format
    - Gently upscale if the image is too small (width < 1600px) to preserve detail
    - Avoid applying extreme contrast or sharpening filters to prevent vision degradation
    """
    if img.mode != "RGB":
        img = img.convert("RGB")

    min_w = 1600
    if img.width < min_w:
        scale = min_w / img.width
        new_h = int(img.height * scale)
        img = img.resize((min_w, new_h), Image.LANCZOS)
        print(f"  [PRE] Upscaled to {min_w}×{new_h}px")

    return img


# ─────────────────────────────────────────────────
#  Gemini call wrapper with retry
# ─────────────────────────────────────────────────
# Primary and fallback models — tried in order when 503 occurs
_PRIMARY_MODEL  = "gemini-2.5-flash"
_FALLBACK_MODEL = "gemini-2.0-flash"

def _call_gemini(img, prompt, retries=4, preprocess=False, config=None):
    """
    Call Gemini with automatic retry and model fallback.
    On 503 (overloaded), waits with exponential back-off then switches
    to _FALLBACK_MODEL before giving up.
    """
    last_error = None
    if preprocess:
        img = _preprocess_for_ocr(img)

    models_to_try = [_PRIMARY_MODEL, _FALLBACK_MODEL]

    for model in models_to_try:
        for attempt in range(1, retries + 2):
            try:
                print(f"  Gemini API [{model}] (attempt {attempt})...")
                resp = client.models.generate_content(
                    model=model,
                    contents=[prompt, img],
                    config=config,
                )
                return resp.text
            except Exception as e:
                last_error = e
                err_str = str(e)
                print(f"  Attempt {attempt} failed: {type(e).__name__}: {e}")
                is_503 = "503" in err_str or "UNAVAILABLE" in err_str
                if attempt <= retries:
                    if is_503:
                        # Exponential backoff: 4s, 8s, 16s, 32s
                        wait = (2 ** attempt) * 2
                        print(f"  High load detected. Retrying in {wait}s...")
                        time.sleep(wait)
                    else:
                        # Non-503 error — no point retrying same model
                        break
                else:
                    if is_503:
                        print(f"  [{model}] still overloaded after {retries+1} attempts.")
                    break  # exhausted retries for this model
        else:
            # All retries exhausted without breaking — move to next model
            pass
        if not ("503" in str(last_error) or "UNAVAILABLE" in str(last_error)):
            # Non-503 failure — no point trying fallback model
            break
        if model != models_to_try[-1]:
            print(f"  Switching to fallback model: {_FALLBACK_MODEL}")
    raise last_error


# ─────────────────────────────────────────────────
#  Allowed spacing values
# ─────────────────────────────────────────────────
ALLOWED_SPACING = (0, 3, 6, 9, 12, 18, 24)

def _snap_spacing(val, font_size_pt=12):
    try:
        val = float(val)
    except (TypeError, ValueError):
        val = 0.0
    return min(ALLOWED_SPACING, key=lambda x: abs(x - val))


# ─────────────────────────────────────────────────
#  OCR text post-processing cleanup
# ─────────────────────────────────────────────────
# Common Gemini OCR confusions: fix recurring misreads
_OCR_FIXES = [
    # Zero / O confusion in obvious numeric contexts
    (re.compile(r'(?<=\d)O(?=\d)'), '0'),   # 1O5 → 105
    (re.compile(r'(?<=\d)o(?=\d)'), '0'),   # 1o5 → 105
    # Trailing/leading whitespace inside parens
    (re.compile(r'\(\s+'), '('),
    (re.compile(r'\s+\)'), ')'),
    # Stray backticks from markdown leakage
    (re.compile(r'`+'), ''),
    # Remove accidental prefix repetition e.g. "BODY: BODY: text"
    (re.compile(r'^(HEADING|SUBHEAD|BODY|BULLET|CENTER|UNDERLN):\s*\1:\s*', re.I), r'\1: '),
]

def _clean_ocr_text(text: str) -> str:
    """Apply post-processing fixes to a single transcribed line."""
    for pattern, replacement in _OCR_FIXES:
        text = pattern.sub(replacement, text)
    
    text = text.rstrip()
    
    # Convert multiple interior spaces to non-breaking spaces to preserve alignment
    # e.g., "x  =  y" -> "x \xA0= \xA0y" (preserves spacing but allows word wrap on single spaces)
    while "  " in text:
        text = text.replace("  ", " \xA0")
        
    # Convert leading spaces to non-breaking spaces so DOCX/HTML don't collapse them
    match = re.match(r'^(\s+)(.*)$', text)
    if match:
        text = match.group(1).replace(" ", "\xA0") + match.group(2)
        
    return text


# ─────────────────────────────────────────────────
#  Parse raw OCR text lines into structured elements
# ─────────────────────────────────────────────────
# Map prefix → (bold, italic, underline, font_size_pt, alignment, left_indent_cm, space_before_pt, space_after_pt)
_PREFIX_MAP = {
    "HEADING":  (True,  False, False, 16, "left",   0.0, 12, 6),
    "SUBHEAD":  (True,  False, False, 13, "left",   0.0,  8, 4),
    "BODY":     (False, False, False, 12, "left",   0.0,  0, 3),
    "BULLET":   (False, False, False, 12, "left",   0.5,  0, 3),
    "CENTER":   (False, False, False, 12, "center", 0.0,  4, 4),
    "UNDERLN":  (False, False, True,  12, "left",   0.0,  0, 3),
    "TABLE":    (False, False, False, 11, "left",   0.0,  4, 4),
}

def _parse_text_lines(raw_text):
    """
    Convert the structured OCR output into a list of element dicts.
    Handles all PREFIX: tags produced by TEXT_OCR_PROMPT.
    Falls back gracefully for un-tagged lines (treated as BODY).
    """
    elements = []
    # Collapse runs of 3+ blank lines into 2 maximum
    collapsed = re.sub(r'\n{3,}', '\n\n', raw_text)

    last_equal_idx = -1

    for line in collapsed.split("\n"):
        line_clean = line.rstrip()

        # True blank line → paragraph gap
        if line_clean.strip() == "":
            last_equal_idx = -1
            # Avoid stacking multiple blank_lines
            if elements and elements[-1].get("type") != "blank_line":
                elements.append({"type": "blank_line"})
            continue

        # Detect and strip prefix smoothly while preserving leading spaces
        prefix = "BODY"
        text = line_clean
        for pfx in _PREFIX_MAP:
            # Match prefix (case-insensitive), an optional colon-space, and capture the rest
            pattern = rf'^\s*{pfx}:\s?(.*)$'
            match = re.match(pattern, line_clean, re.IGNORECASE)
            if match:
                prefix = pfx
                text = match.group(1)  # Preserves leading spaces of the actual content
                break

        # Clean the transcribed text
        text = _clean_ocr_text(text)

        if not text:
            continue

        if prefix == "TABLE":
            last_equal_idx = -1
            # Parse markdown row: | col1 | col2 |
            # Clean up the \xA0 spaces specifically for table cells
            clean_row = text.replace('\xA0', ' ').strip()
            cells = [c.strip() for c in clean_row.strip('|').split('|')]
            
            # Check if it's a markdown separator row (e.g., |---|---| )
            is_separator = all(re.match(r'^[-:]+$', c) for c in cells if c)
            
            if not is_separator:
                if elements and elements[-1].get("type") == "table":
                    elements[-1]["data"].append(cells)
                else:
                    elements.append({
                        "type": "table",
                        "data": [cells]
                    })
            continue

        # ── Smart Math Equation '=' Alignment ──
        # If the line starts with '=' (ignoring spaces), pad it so the '='
        # aligns horizontally with the '=' in the previous equation line.
        temp_text = text.replace("\xA0", " ")
        first_char_idx = len(temp_text) - len(temp_text.lstrip())
        
        if first_char_idx < len(temp_text) and temp_text[first_char_idx] == '=' and last_equal_idx > 0:
            # It's a continuation line starting with =
            needed_spaces = last_equal_idx - first_char_idx
            if needed_spaces > 0:
                text = ("\xA0" * needed_spaces) + text
        else:
            # Look for an '=' in a normal line to anchor future continuations
            eq_pos = temp_text.find("=")
            if eq_pos > 0:
                last_equal_idx = eq_pos

        bold, italic, underline, fsize, align, left_i, sp_bef, sp_aft = _PREFIX_MAP[prefix]

        # Add space_before for headings only if previous element was not blank
        if prefix in ("HEADING", "SUBHEAD") and elements and elements[-1].get("type") != "blank_line":
            sp_bef = max(sp_bef, 10)

        elements.append({
            "type":                 "text",
            "text":                 text,
            "alignment":            align,
            "font_size_pt":         fsize,
            "bold":                 bold,
            "italic":               italic,
            "underline":            underline,
            "first_line_indent_cm": 0.0,
            "left_indent_cm":       left_i,
            "space_before_pt":      sp_bef,
            "space_after_pt":       sp_aft,
        })
    return elements


# ─────────────────────────────────────────────────
#  Arrow Unicode mapping
# ─────────────────────────────────────────────────
_ARROW_UNICODE = {
    "right":      "⟶",   # Long rightwards arrow
    "left":       "⟵",   # Long leftwards arrow
    "up":         "↑",    # Upwards arrow
    "down":       "↓",    # Downwards arrow
    "up-right":   "↗",   # North-east arrow
    "up-left":    "↖",   # North-west arrow
    "down-right": "↘",   # South-east arrow
    "down-left":  "↙",   # South-west arrow
}

def _resolve_arrow_char(direction: str) -> str:
    """Return the best Unicode arrow for the given direction string."""
    d = (direction or "").strip().lower()
    return _ARROW_UNICODE.get(d, "⟶")   # default: rightwards


# ─────────────────────────────────────────────────
#  Bracket / Brace Unicode mapping
# ─────────────────────────────────────────────────
# Maps (style, side) -> (display_char, font_size_multiplier)
_BRACKET_CHARS = {
    # curly braces
    ("curly",  "left"):  "⎧⎨⎩",   # top, mid, bottom curly bracket pieces
    ("curly",  "right"): "⎫⎬⎭",
    # square brackets
    ("square", "left"):  "[",
    ("square", "right"): "]",
    # plain / L-bracket
    ("plain",  "left"):  "⎡⎣",   # top and bottom square bracket pieces
    ("plain",  "right"): "⎤⎦",
}

# Single display chars for simple cases (used for PDF/DOCX single-char rendering)
_BRACKET_SINGLE = {
    ("curly",  "left"):  "{",
    ("curly",  "right"): "}",
    ("square", "left"):  "[",
    ("square", "right"): "]",
    ("plain",  "left"):  "[",
    ("plain",  "right"): "]",
}

def _resolve_bracket_char(style: str, side: str) -> str:
    """Return the best single bracket character for style+side."""
    s = (style or "square").strip().lower()
    d = (side  or "left").strip().lower()
    return _BRACKET_SINGLE.get((s, d), "[")


# ─────────────────────────────────────────────────
#  Merge text elements with drawing inserts
#  Drawings are inserted at the correct Y position
#  in the text flow using position_y_fraction.
# ─────────────────────────────────────────────────
def _merge_elements(text_elements, drawings):
    """
    Insert drawing elements into the text flow at the position that matches
    their vertical location in the source image.

    Simple-arrow drawings become type="arrow" (rendered as Unicode text).
    Real drawings become type="drawing" (rendered as inline cropped image).
    """
    if not drawings:
        return text_elements

    text_only = [e for e in text_elements if e.get("type") == "text"]
    total = max(1, len(text_only))

    insertions = []
    for d in sorted(drawings, key=lambda x: x.get("position_y_fraction", 0.5)):
        frac = float(d.get("position_y_fraction", 0.5))
        idx  = int(frac * total)
        if d.get("is_simple_arrow", False):
            # Represent as a Unicode arrow text element
            arrow_char = _resolve_arrow_char(d.get("arrow_direction", "right"))
            insertions.append((idx, {
                "type":        "arrow",
                "arrow_char":  arrow_char,
                "description": d.get("description", ""),
            }))
            print(f"  [MERGE] Arrow-only drawing -> Unicode '{arrow_char}'")
        elif d.get("is_simple_bracket", False):
            # Represent as a Unicode bracket character
            brk_char = _resolve_bracket_char(
                d.get("bracket_style", "square"),
                d.get("bracket_side",  "left"),
            )
            insertions.append((idx, {
                "type":          "bracket",
                "bracket_char":  brk_char,
                "bracket_style": d.get("bracket_style", "square"),
                "bracket_side":  d.get("bracket_side",  "left"),
                "description":   d.get("description", ""),
            }))
            print(f"  [MERGE] Bracket-only drawing -> Unicode '{brk_char}'")
        else:
            insertions.append((idx, {
                "type":        "drawing",
                "bbox":        d.get("bbox"),
                "description": d.get("description", ""),
            }))

    # Walk through text_elements, mapping their text-only index to a position
    result = []
    text_idx = 0          # counts only "text" type elements seen so far
    ins_ptr  = 0          # pointer into insertions list

    for el in text_elements:
        # Before adding this element, check if any drawing should be inserted
        while ins_ptr < len(insertions) and insertions[ins_ptr][0] <= text_idx:
            _, drw = insertions[ins_ptr]
            result.append({
                "type": "drawing",
                "bbox": drw.get("bbox"),
                "description": drw.get("description", ""),
            })
            ins_ptr += 1

        result.append(el)
        if el.get("type") == "text":
            text_idx += 1

    # Append any remaining drawings at the end
    while ins_ptr < len(insertions):
        _, drw = insertions[ins_ptr]
        result.append({
            "type": "drawing",
            "bbox": drw.get("bbox"),
            "description": drw.get("description", ""),
        })
        ins_ptr += 1

    return result


# ─────────────────────────────────────────────────
#  Auto OCR — two-step pipeline
#  Step 1: pure text OCR (high accuracy)
#  Step 2: layout/drawing detection
#  Then merge.
# ─────────────────────────────────────────────────
def run_ocr_auto(image_path):
    print(f"[AUTO] Consolidated OCR & Layout: {image_path}")
    raw_img = Image.open(image_path)

    # Standardize image (convert to RGB, upscale if too small)
    img = _preprocess_for_ocr(raw_img)

    try:
        print("  Calling Gemini for combined layout and text OCR...")
        config = genai_types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=DocumentLayout,
            temperature=0.1,
        )
        
        resp_text = _call_gemini(img, COMBINED_OCR_PROMPT, config=config)
        layout_data = DocumentLayout.model_validate_json(resp_text)
        print(f"  Gemini Layout: margin={layout_data.page_margin_cm}cm, spacing={layout_data.line_spacing}")
        
        elements = []
        for el in layout_data.elements:
            etype = el.type
            if etype == "blank_line":
                elements.append({"type": "blank_line"})
            elif etype == "table":
                elements.append({
                    "type": "table",
                    "data": el.table_data or []
                })
            elif etype == "drawing":
                if el.is_simple_arrow:
                    arrow_char = _resolve_arrow_char(el.arrow_direction or "right")
                    elements.append({
                        "type":        "arrow",
                        "arrow_char":  arrow_char,
                        "description": el.description or "",
                    })
                elif el.is_simple_bracket:
                    brk_char = _resolve_bracket_char(
                        el.bracket_style or "square",
                        el.bracket_side or "left"
                    )
                    elements.append({
                        "type":          "bracket",
                        "bracket_char":  brk_char,
                        "bracket_style": el.bracket_style or "square",
                        "bracket_side":  el.bracket_side or "left",
                        "description":   el.description or "",
                    })
                else:
                    elements.append({
                        "type":        "drawing",
                        "bbox":        el.bbox,
                        "description": el.description or "",
                    })
            elif etype == "text":
                tag = (el.tag or "BODY").upper()
                text = _clean_ocr_text(el.text or "")
                if not text:
                    continue
                
                # Fetch settings from map
                bold_mapped, italic_mapped, underline_mapped, fsize, align, left_i, sp_bef, sp_aft = _PREFIX_MAP.get(
                    tag, _PREFIX_MAP["BODY"]
                )
                
                # Use model-detected override values if specified, else fall back to prefix defaults
                bold = el.bold if el.bold is not None else bold_mapped
                italic = el.italic if el.italic is not None else italic_mapped
                underline = el.underline if el.underline is not None else underline_mapped
                align = el.alignment if el.alignment is not None else align
                left_i = el.left_indent_cm if el.left_indent_cm is not None else left_i
                
                elements.append({
                    "type":                 "text",
                    "text":                 text,
                    "alignment":            align,
                    "font_size_pt":         fsize,
                    "bold":                 bold,
                    "italic":               italic,
                    "underline":            underline,
                    "first_line_indent_cm": 0.0,
                    "left_indent_cm":       left_i,
                    "space_before_pt":      sp_bef,
                    "space_after_pt":       sp_aft,
                })
        
        meta = {
            "line_spacing": layout_data.line_spacing,
            "page_margin_cm": layout_data.page_margin_cm
        }
        return elements, meta

    except Exception as e:
        print(f"  Combined OCR failed ({e}), trying plain OCR fallback...")
        try:
            raw_text = _call_gemini(img, PLAIN_OCR_PROMPT)
            text_elements = _parse_text_lines(raw_text)
            meta = {"line_spacing": 1.15, "page_margin_cm": 2.54}
            return text_elements, meta
        except Exception as e2:
            err = f"[OCR FAILED: {type(e2).__name__}: {str(e2)[:200]}]"
            return [{"type": "text", "text": err, "alignment": "left",
                     "font_size_pt": 12, "bold": False, "italic": False,
                     "underline": False, "first_line_indent_cm": 0.0,
                     "left_indent_cm": 0.0, "space_before_pt": 0, "space_after_pt": 0}], \
                   {"line_spacing": 1.15, "page_margin_cm": 2.54}


# ─────────────────────────────────────────────────
#  Manual OCR — plain text only
# ─────────────────────────────────────────────────
def run_ocr_manual(image_path):
    print(f"[MANUAL] Plain OCR: {image_path}")
    img = Image.open(image_path)
    try:
        raw = _call_gemini(img, PLAIN_OCR_PROMPT)
        return [l.strip() for l in raw.split("\n") if l.strip()]
    except Exception as e:
        return [f"[OCR FAILED: {type(e).__name__}: {str(e)[:200]}]"]


# ─────────────────────────────────────────────────
#  Drawing crop helper
# ─────────────────────────────────────────────────
def _crop_drawing(image_path, bbox):
    """
    Crop the drawing region from image_path using normalised bbox [x1,y1,x2,y2].
    Returns the path to a cropped temp PNG, or None on any failure.
    """
    try:
        if not bbox or len(bbox) != 4:
            return None
        x1n, y1n, x2n, y2n = [float(v) for v in bbox]
        x1n = max(0.0, min(1.0, x1n))
        y1n = max(0.0, min(1.0, y1n))
        x2n = max(0.0, min(1.0, x2n))
        y2n = max(0.0, min(1.0, y2n))
        if x2n <= x1n or y2n <= y1n:
            return None
        img = Image.open(image_path)
        w, h = img.size
        cropped = img.crop((int(x1n*w), int(y1n*h), int(x2n*w), int(y2n*h)))
        out_path = f"{UPLOAD_DIR}/{uuid.uuid4()}_crop.png"
        cropped.save(out_path)
        print(f"  Cropped drawing: {out_path} ({int((x2n-x1n)*w)}×{int((y2n-y1n)*h)}px)")
        return out_path
    except Exception as e:
        print(f"  _crop_drawing failed: {e}")
        return None


# ─────────────────────────────────────────────────
#  Docx helpers
# ─────────────────────────────────────────────────
def _apply_run(run, font_family, font_size_pt, bold=False, italic=False, underline=False):
    run.font.name      = font_family
    run.font.size      = Pt(font_size_pt)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.underline = underline


def _apply_para(para, alignment_str, line_spacing,
                space_before_pt=0, space_after_pt=0,
                first_line_indent_cm=0.0, left_indent_cm=0.0):
    fmt = para.paragraph_format
    fmt.alignment         = ALIGN_MAP.get(alignment_str, WD_ALIGN_PARAGRAPH.LEFT)
    fmt.line_spacing      = line_spacing
    fmt.space_before      = Pt(space_before_pt)
    fmt.space_after       = Pt(space_after_pt)
    fmt.first_line_indent = Cm(first_line_indent_cm)
    fmt.left_indent       = Cm(left_indent_cm)


def _set_margins(doc, margin_cm):
    margin_cm = max(1.0, min(5.0, float(margin_cm)))
    for section in doc.sections:
        section.top_margin    = Cm(margin_cm)
        section.bottom_margin = Cm(margin_cm)
        section.left_margin   = Cm(margin_cm)
        section.right_margin  = Cm(margin_cm)


def _add_inline_picture_docx(doc, img_path, bbox, page_margin_cm=2.54):
    """
    Add a drawing INLINE in the text flow (not floating).
    Width is capped to the usable content width. Aspect ratio is preserved.
    This approach guarantees zero text-image overlap.
    """
    cropped = _crop_drawing(img_path, bbox)
    if not cropped:
        return

    try:
        # Calculate maximum allowed width (content area width)
        content_w_cm = 21.0 - 2.0 * page_margin_cm
        max_w_cm = max(2.0, content_w_cm)

        # Get natural bbox dimensions as a fraction of content width
        x1n, y1n, x2n, y2n = [float(v) for v in bbox]
        bbox_w_frac = x2n - x1n   # fraction of full width
        bbox_h_frac = y2n - y1n

        # Map to cm, cap at content width
        img_w_cm = min(max_w_cm, bbox_w_frac * (21.0 - 2.0 * page_margin_cm))
        img_w_cm = max(1.0, img_w_cm)

        # Calculate height from aspect ratio
        with Image.open(cropped) as ci:
            nat_w, nat_h = ci.size
        aspect = nat_h / nat_w if nat_w > 0 else 0.5
        img_h_cm = img_w_cm * aspect

        # Clamp height so very tall drawings don't overflow page
        max_h_cm = 29.7 - 2.0 * page_margin_cm
        if img_h_cm > max_h_cm * 0.7:
            img_h_cm = max_h_cm * 0.7
            img_w_cm = img_h_cm / aspect if aspect > 0 else img_w_cm

        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after  = Pt(6)
        para.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(cropped, width=Cm(img_w_cm), height=Cm(img_h_cm))
        print(f"  [DOCX] Drawing inline: {img_w_cm:.2f}×{img_h_cm:.2f}cm")
    except Exception as e:
        print(f"  [DOCX] Inline image failed: {e}")


# ─────────────────────────────────────────────────
#  PDF helpers
# ─────────────────────────────────────────────────
PT_PER_CM = 28.3465

def _rl_font(bold, italic):
    if bold and italic: return "Helvetica-BoldOblique"
    if bold:            return "Helvetica-Bold"
    if italic:          return "Helvetica-Oblique"
    return "Helvetica"


def _pdf_draw_text(c, txt, x, y, font_name, font_size, underline=False):
    c.setFont(font_name, font_size)
    c.drawString(x, y, txt)
    if underline:
        tw = c.stringWidth(txt, font_name, font_size)
        c.setLineWidth(0.5)
        c.line(x, y - 1.5, x + tw, y - 1.5)


def _pdf_draw_table(c, table_data, y, margin_pts, page_w, page_h):
    """
    Draw a crisp native table on the PDF canvas.
    Returns the new y position.
    """
    if not table_data:
        return y

    rows = len(table_data)
    cols = max((len(r) for r in table_data), default=0)
    if cols == 0:
        return y

    # Calculate column widths using stringWidth
    col_widths = [0] * cols
    pad_x = 4
    pad_y = 4
    fsize = 11
    font_name = "Helvetica"
    font_bold = "Helvetica-Bold"

    for r_idx, r in enumerate(table_data):
        f = font_bold if r_idx == 0 else font_name
        for i, cell in enumerate(r):
            if i < cols:
                w = c.stringWidth(cell, f, fsize)
                col_widths[i] = max(col_widths[i], w + pad_x * 2)

    usable_w = page_w - 2 * margin_pts
    total_w = sum(col_widths)
    
    # Scale or stretch widths to fit page logically
    if total_w < usable_w * 0.5:
        # Stretch to 50% minimum
        scale = (usable_w * 0.5) / total_w
        col_widths = [cw * scale for cw in col_widths]
    elif total_w > usable_w:
        # Squeeze to fit page
        scale = usable_w / total_w
        col_widths = [cw * scale for cw in col_widths]
        
    x_start = margin_pts
    row_h = fsize + pad_y * 2
    
    y -= row_h  # start position for first row bottom
    
    # Simple page wrap check
    if y - (rows * row_h) < margin_pts and rows < 30:
        c.showPage()
        y = page_h - margin_pts - row_h

    # Draw grid and text
    c.setLineWidth(0.5)
    for row_idx, r in enumerate(table_data):
        curr_x = x_start
        for col_idx in range(cols):
            cell_txt = r[col_idx] if col_idx < len(r) else ""
            cw = col_widths[col_idx]
            
            # Draw cell box
            c.rect(curr_x, y, cw, row_h)
            
            # Draw text
            f = font_bold if row_idx == 0 else font_name
            c.setFont(f, fsize)
            text_x = curr_x + pad_x
            text_y = y + pad_y + 1
            c.drawString(text_x, text_y, cell_txt)
            
            curr_x += cw
            
        y -= row_h
        if y < margin_pts:
            c.showPage()
            y = page_h - margin_pts - row_h
            
    return y - 10

def _pdf_place_drawing_inline(c, image_path, bbox, y, page_w, margin_pts, page_h, line_spacing=1.15):
    """
    Place a drawing INLINE in the PDF text flow at the current y position.
    Returns the height consumed (in points) so the caller can advance y correctly.
    """
    if not bbox or len(bbox) != 4:
        return 0

    cropped = _crop_drawing(image_path, bbox)
    if not cropped:
        return 0

    try:
        content_w = page_w - 2.0 * margin_pts

        x1n, y1n, x2n, y2n = [float(v) for v in bbox]
        bbox_w_frac = x2n - x1n
        bbox_h_frac = y2n - y1n

        # Natural pixel dimensions → aspect ratio
        with Image.open(cropped) as ci:
            nat_w, nat_h = ci.size
        aspect = nat_h / nat_w if nat_w > 0 else 0.5

        # Scale width to fit content area
        draw_w = min(content_w, max(50, bbox_w_frac * content_w))
        draw_h = draw_w * aspect

        # Cap height at 60% of usable page height
        max_h = (page_h - 2.0 * margin_pts) * 0.6
        if draw_h > max_h:
            draw_h = max_h
            draw_w = draw_h / aspect if aspect > 0 else draw_w

        # Centre the drawing horizontally
        x = margin_pts + (content_w - draw_w) / 2.0

        # Check if there's enough room on the current page; if not, start new page
        if y - draw_h < margin_pts:
            c.showPage()
            y = page_h - margin_pts

        # ReportLab y is from bottom; we track y from top
        y_bottom = y - draw_h

        ir = ImageReader(cropped)
        c.drawImage(ir, x, y_bottom, width=draw_w, height=draw_h,
                    preserveAspectRatio=False, mask='auto')

        print(f"  [PDF] Inline drawing: x={x:.1f} y_btm={y_bottom:.1f} {draw_w:.1f}×{draw_h:.1f}pts")
        return draw_h + 12   # height consumed + bottom padding
    except Exception as e:
        print(f"  [PDF] Inline drawing failed: {e}")
        return 0


# ═══════════════════════════════════════════════════
#  CONVERT ENDPOINT
# ═══════════════════════════════════════════════════
@app.post("/convert")
async def convert(
    file:              UploadFile = File(...),
    mode:              str   = Form("preserve"),
    outtype:           str   = Form("docx"),
    auto_format:       str   = Form("true"),
    # Manual params
    font_family:       str   = Form("Calibri"),
    font_size:         float = Form(12.0),
    line_spacing:      float = Form(1.15),
    para_spacing:      float = Form(8.0),
    first_line_indent: float = Form(0.0),
    page_margin:       float = Form(2.54),
    text_align:        str   = Form("left"),
    text_bold:         str   = Form("false"),
    text_italic:       str   = Form("false"),
    text_underline:    str   = Form("false"),
):
    uid        = str(uuid.uuid4())
    input_path = f"{UPLOAD_DIR}/{uid}_{file.filename}"

    with open(input_path, "wb") as f:
        f.write(await file.read())

    use_auto = auto_format.lower() == "true"

    if input_path.lower().endswith(".pdf"):
        pages = pdf_to_images(input_path)
    else:
        pages = [input_path]

    # ═══════════════════════════
    #  WORD (.docx)
    # ═══════════════════════════
    if outtype == "docx":
        doc = Document()

        if use_auto:
            all_data = [run_ocr_auto(p) for p in pages]
            _set_margins(doc, all_data[0][1].get("page_margin_cm", 2.54))

            for i, (elements, meta) in enumerate(all_data):
                if i > 0:
                    doc.add_page_break()
                ls = meta.get("line_spacing", 1.15)
                pm = meta.get("page_margin_cm", 2.54)

                for el in elements:
                    etype = el.get("type", "text")

                    if etype == "blank_line":
                        para = doc.add_paragraph()
                        para.paragraph_format.space_after = Pt(4)

                    elif etype == "drawing":
                        # Inline image — sits naturally in the text flow
                        _add_inline_picture_docx(
                            doc,
                            img_path=pages[i],
                            bbox=el.get("bbox"),
                            page_margin_cm=pm,
                        )

                    elif etype == "arrow":
                        # Simple arrow — render as a large Unicode character, no image
                        arrow_char = el.get("arrow_char", "⟶")
                        para = doc.add_paragraph()
                        para.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER
                        para.paragraph_format.space_before = Pt(6)
                        para.paragraph_format.space_after  = Pt(6)
                        run = para.add_run(arrow_char)
                        run.font.name = "Segoe UI Symbol"
                        run.font.size = Pt(24)
                        run.font.bold = False
                        print(f"  [DOCX] Arrow element: '{arrow_char}'")

                    elif etype == "bracket":
                        # Grouping bracket — render as a tall bold Unicode char, no image
                        brk_char  = el.get("bracket_char", "[")
                        brk_side  = el.get("bracket_side", "left")
                        brk_style = el.get("bracket_style", "square")
                        # Choose alignment: left-bracket aligns left, right-bracket aligns right
                        align_val = (WD_ALIGN_PARAGRAPH.LEFT
                                     if brk_side == "left"
                                     else WD_ALIGN_PARAGRAPH.RIGHT)
                        para = doc.add_paragraph()
                        para.paragraph_format.alignment    = align_val
                        para.paragraph_format.space_before = Pt(4)
                        para.paragraph_format.space_after  = Pt(4)
                        run = para.add_run(brk_char)
                        # Use a larger font so the bracket is visually tall
                        run.font.name = "Segoe UI Symbol"
                        run.font.size = Pt(36)
                        run.font.bold = (brk_style == "plain")  # bold for plain brackets
                        print(f"  [DOCX] Bracket element: '{brk_char}' ({brk_style}, {brk_side})")

                    elif etype == "table":
                        # Native DOCX Table
                        table_data = el.get("data", [])
                        if table_data:
                            rows = len(table_data)
                            cols = max((len(r) for r in table_data), default=0)
                            if cols > 0:
                                table = doc.add_table(rows=rows, cols=cols)
                                table.style = 'Table Grid'
                                
                                for row_idx, row_data in enumerate(table_data):
                                    for col_idx, cell_text in enumerate(row_data):
                                        if col_idx < cols:
                                            cell = table.cell(row_idx, col_idx)
                                            cell.text = cell_text
                                            for paragraph in cell.paragraphs:
                                                for run in paragraph.runs:
                                                    run.font.name = "Calibri"
                                                    run.font.size = Pt(11)
                                                    # Bold first row as header heuristic
                                                    if row_idx == 0:
                                                        run.font.bold = True
                                
                                # Add space after table
                                doc.add_paragraph().paragraph_format.space_after = Pt(8)
                                print(f"  [DOCX] Table element: {rows}x{cols}")

                    else:  # "text"
                        para = doc.add_paragraph()
                        run  = para.add_run(el.get("text", ""))
                        _apply_run(
                            run, "Calibri",
                            el.get("font_size_pt", 12),
                            el.get("bold", False),
                            el.get("italic", False),
                            el.get("underline", False),
                        )
                        _apply_para(
                            para,
                            el.get("alignment", "left"),
                            ls,
                            el.get("space_before_pt", 0),
                            el.get("space_after_pt", 4),
                            el.get("first_line_indent_cm", 0.0),
                            el.get("left_indent_cm", 0.0),
                        )

        else:
            bold      = text_bold.lower()      == "true"
            italic    = text_italic.lower()    == "true"
            underline = text_underline.lower() == "true"
            _set_margins(doc, page_margin)

            for i, page in enumerate(pages):
                lines = run_ocr_manual(page)
                if mode == "preserve" and i > 0:
                    doc.add_page_break()
                if mode == "flow":
                    text = " ".join(lines)
                    if text.strip():
                        para = doc.add_paragraph()
                        run  = para.add_run(text)
                        _apply_run(run, font_family, font_size, bold, italic, underline)
                        _apply_para(para, text_align, line_spacing,
                                    0, _snap_spacing(para_spacing), first_line_indent)
                else:
                    for line in lines:
                        para = doc.add_paragraph()
                        run  = para.add_run(line)
                        _apply_run(run, font_family, font_size, bold, italic, underline)
                        _apply_para(para, text_align, line_spacing,
                                    0, _snap_spacing(para_spacing), first_line_indent)

        out_path = f"{OUTPUT_DIR}/{uid}.docx"
        doc.save(out_path)
        return FileResponse(out_path, filename="output.docx")

    # ═══════════════════════════
    #  PDF
    # ═══════════════════════════
    else:
        out_path = f"{OUTPUT_DIR}/{uid}.pdf"
        c = canvas.Canvas(out_path, pagesize=A4)
        page_w, page_h = A4

        if use_auto:
            for pg_path in pages:
                elements, meta = run_ocr_auto(pg_path)
                margin_pts = max(28, meta.get("page_margin_cm", 2.54) * PT_PER_CM)
                usable_w   = page_w - 2 * margin_pts
                ls         = meta.get("line_spacing", 1.15)

                # Single pass: render text AND drawings in order, top to bottom
                y = page_h - margin_pts

                for el in elements:
                    etype = el.get("type", "text")

                    if etype == "blank_line":
                        y -= 8
                        if y < margin_pts:
                            c.showPage()
                            y = page_h - margin_pts
                        continue

                    if etype == "arrow":
                        # Simple arrow — render as large centred Unicode character
                        arrow_char  = el.get("arrow_char", "⟶")
                        fsize_arrow = 24
                        y -= 6
                        if y < margin_pts:
                            c.showPage()
                            y = page_h - margin_pts
                        tw = c.stringWidth(arrow_char, "Helvetica", fsize_arrow)
                        x_arrow = margin_pts + (usable_w - tw) / 2
                        c.setFont("Helvetica", fsize_arrow)
                        c.drawString(x_arrow, y, arrow_char)
                        y -= fsize_arrow + 6
                        print(f"  [PDF] Arrow element: '{arrow_char}'")
                        continue

                    if etype == "bracket":
                        # Grouping bracket — render as tall Unicode character in the flow
                        brk_char  = el.get("bracket_char", "[")
                        brk_side  = el.get("bracket_side", "left")
                        fsize_brk = 40   # tall bracket
                        y -= 4
                        if y < margin_pts:
                            c.showPage()
                            y = page_h - margin_pts
                        tw = c.stringWidth(brk_char, "Helvetica-Bold", fsize_brk)
                        # Left bracket aligns left, right bracket aligns right
                        if brk_side == "right":
                            x_brk = margin_pts + usable_w - tw
                        else:
                            x_brk = margin_pts
                        c.setFont("Helvetica-Bold", fsize_brk)
                        c.drawString(x_brk, y, brk_char)
                        y -= fsize_brk + 4
                        print(f"  [PDF] Bracket element: '{brk_char}' ({brk_side})")
                        continue

                    if etype == "table":
                        table_data = el.get("data", [])
                        y -= 6
                        y = _pdf_draw_table(c, table_data, y, margin_pts, page_w, page_h)
                        print(f"  [PDF] Table element")
                        continue

                    if etype == "drawing":
                        # Place drawing inline at current y, advance y by height consumed
                        consumed = _pdf_place_drawing_inline(
                            c, pg_path, el.get("bbox"), y,
                            page_w, margin_pts, page_h, ls
                        )
                        y -= consumed
                        continue

                    # text element
                    txt       = el.get("text", "")
                    fsize     = float(el.get("font_size_pt", 12))
                    bold      = el.get("bold", False)
                    italic    = el.get("italic", False)
                    underline = el.get("underline", False)
                    align     = el.get("alignment", "left")
                    indent    = el.get("first_line_indent_cm", 0.0) * PT_PER_CM
                    left_i    = el.get("left_indent_cm", 0.0) * PT_PER_CM
                    sp_before = float(el.get("space_before_pt", 0))
                    sp_after  = float(el.get("space_after_pt", 4))
                    leading   = fsize * ls

                    font_name = _rl_font(bold, italic)

                    y -= sp_before
                    if y < margin_pts:
                        c.showPage()
                        y = page_h - margin_pts

                    x = margin_pts + indent + left_i
                    if align == "center":
                        tw = c.stringWidth(txt, font_name, fsize)
                        x  = margin_pts + (usable_w - tw) / 2
                    elif align == "right":
                        tw = c.stringWidth(txt, font_name, fsize)
                        x  = margin_pts + usable_w - tw

                    _pdf_draw_text(c, txt, x, y, font_name, fsize, underline)
                    y -= leading + sp_after

                    if y < margin_pts:
                        c.showPage()
                        y = page_h - margin_pts

                c.showPage()

        else:
            bold       = text_bold.lower()      == "true"
            italic     = text_italic.lower()    == "true"
            underline  = text_underline.lower() == "true"
            margin_pts = float(page_margin) * PT_PER_CM
            usable_w   = page_w - 2 * margin_pts
            leading    = float(font_size) * float(line_spacing)
            indent_pts = float(first_line_indent) * PT_PER_CM
            font_name  = _rl_font(bold, italic)
            sp_after   = _snap_spacing(para_spacing)

            for pg_path in pages:
                lines = run_ocr_manual(pg_path)
                y     = page_h - margin_pts

                if mode == "flow":
                    text  = " ".join(lines)
                    words = text.split(" ")
                    line_str = ""
                    for w in words:
                        test = line_str + w + " "
                        if c.stringWidth(test, font_name, font_size) > usable_w:
                            _pdf_draw_text(c, line_str.strip(), margin_pts, y, font_name, font_size, underline)
                            y -= leading
                            if y < margin_pts:
                                c.showPage(); y = page_h - margin_pts
                            line_str = w + " "
                        else:
                            line_str = test
                    if line_str.strip():
                        _pdf_draw_text(c, line_str.strip(), margin_pts, y, font_name, font_size, underline)
                        y -= leading + sp_after
                else:
                    for line in lines:
                        x = margin_pts + indent_pts
                        if text_align == "center":
                            tw = c.stringWidth(line, font_name, font_size)
                            x  = margin_pts + (usable_w - tw) / 2
                        elif text_align == "right":
                            tw = c.stringWidth(line, font_name, font_size)
                            x  = margin_pts + usable_w - tw
                        if y < margin_pts:
                            c.showPage(); y = page_h - margin_pts
                        _pdf_draw_text(c, line, x, y, font_name, font_size, underline)
                        y -= leading + sp_after * 0.75

                c.showPage()

        c.save()
        return FileResponse(out_path, filename="output.pdf")