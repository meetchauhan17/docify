<div align="center">

# Docify AI

**High-fidelity handwriting-to-document conversion powered by Google Gemini and multi-provider vision models**

[![Python Version](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.110%2B-009688?style=flat-square&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![Gemini](https://img.shields.io/badge/Gemini-2.5%20Flash-4285F4?style=flat-square&logo=google&logoColor=white)](https://ai.google.dev/)
[![Groq](https://img.shields.io/badge/Groq-Llama%204%20Scout-F55036?style=flat-square)](https://groq.com/)
[![OpenRouter](https://img.shields.io/badge/OpenRouter-Free%20Vision-7E22CE?style=flat-square)](https://openrouter.ai/)
[![ReportLab](https://img.shields.io/badge/ReportLab-PDF%20Engine-E74C3C?style=flat-square)](https://www.reportlab.com/)
[![python-docx](https://img.shields.io/badge/python--docx-DOCX%20Engine-2E86AB?style=flat-square)](https://python-docx.readthedocs.io/)
[![Code Style](https://img.shields.io/badge/Code%20Style-Black-000000?style=flat-square)](https://github.com/psf/black)
[![Platform Support](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)](#installation-guide)
[![License](https://img.shields.io/badge/License-MIT-8E44AD?style=flat-square)](LICENSE)

Docify AI converts scanned handwritten notes, photographic records, and multi-page PDF files into highly structured, fully editable Microsoft Word (.docx) and PDF (.pdf) documents. By executing layout analysis and OCR inside unified multimodal visual prompts, the application reproduces paragraph styles, font hierarchies, mathematical equations, data tables, and inline drawing sketches with optimal structural preservation.

</div>

---

## Table of Contents

- [System Architecture](#system-architecture)
- [Core Capabilities](#core-capabilities)
- [Multi-Provider Resilience Framework](#multi-provider-resilience-framework)
- [Image Preprocessing and OCR Caching](#image-preprocessing-and-ocr-caching)
- [Pydantic Layout Schema](#pydantic-layout-schema)
- [Element Mapping and Output Assembly](#element-mapping-and-output-assembly)
- [Technology Stack](#technology-stack)
- [Directory Structure](#directory-structure)
- [Installation Guide](#installation-guide)
  - [Prerequisites](#prerequisites)
  - [Installing Poppler](#installing-poppler)
  - [Step-by-Step Installation](#step-by-step-installation)
- [Running the Application](#running-the-application)
- [Web Interface Reference](#web-interface-reference)
- [API Reference](#api-reference)
  - [POST /convert](#post-convert)
- [Configuration Reference](#configuration-reference)
- [License](#license)

---

## System Architecture

The following diagram illustrates the lifecycle of a document processing request from file upload to final artifact download, highlighting the pre-processing steps, API provider failover tree, elements assembly, and format rendering.

```mermaid
flowchart TD
    A([User Uploads Document\nImage or PDF]) --> B{File Type?}
    B -- PDF --> C[pdf2image\nExtract per-page PNGs]
    B -- Image --> D[Load Image via Pillow]
    C --> E[Image Preprocessing\nRGB conversion\nLanczos upscaling if width < 1600px]
    D --> E

    E --> F[Master API Dispatcher\n_call_vision_api]
    
    F --> G{Gemini Keys Available?}
    G -- Yes --> H[Call Gemini API\ngemini-2.5-flash / 2.0-flash]
    H --> I{Success?}
    I -- Yes --> V[DocumentLayout JSON]
    I -- No: Quota 429 --> J[Rotate Gemini Key\nRetry up to key count]
    J --> H
    I -- No: Overload 503 --> K[Exponential Backoff\nSwitch to gemini-2.0-flash]
    K --> H
    
    G -- No / Gemini Exhausted --> L{Groq API Key Configured?}
    L -- Yes --> M[Call Groq Vision API\nllama-4-scout-17b / llama-3.2-90b]
    M --> N{Success?}
    N -- Yes --> V
    N -- No --> O[Try Next Groq Model / Fail]
    
    L -- No / Groq Exhausted --> P{OpenRouter Configured?}
    P -- Yes --> Q[Call OpenRouter Vision API\nqwen2.5-vl-72b / llama-3.2-11b]
    Q --> R{Success?}
    R -- Yes --> V
    R -- No --> S[Try Next OpenRouter Model / Fail]
    
    P -- No / All Exhausted --> T[Plain OCR Fallback\nPLAIN_OCR_PROMPT]
    T --> U[Line-by-Line Regex Parser\n_parse_text_lines]
    U --> W[Mapped Elements List]

    V --> W
    
    W --> X[Element Mapping\nrun_ocr_auto]
    X --> Y{Element Type?}
    Y -- text --> Z[Apply Typography Styles\nHEADING / SUBHEAD / BODY / BULLET / CENTER / UNDERLN]
    Y -- table --> AA[Verify Cell Dimensions\nResolve Rowspans & Colspans]
    Y -- drawing --> AB{Visual Shape?}
    AB -- Arrow --> AC[Render Segoe UI Symbol\nUnicode arrow glyph]
    AB -- Bracket --> AD[Render Segoe UI Symbol\nUnicode bracket glyph]
    AB -- Sketch --> AE[_crop_drawing\nCrop bounding box to PNG]
    Y -- blank_line --> AF[Insert Paragraph Spacer]
    
    Z --> AG[Document Assembly Engine]
    AA --> AG
    AC --> AG
    AD --> AG
    AE --> AG
    AF --> AG
    
    AG --> AH{Selected Format?}
    AH -- DOCX --> AI[python-docx Generation\nTables, Runs, Inline Pics]
    AH -- PDF --> AJ[ReportLab Canvas Generation\nWrapped Text, Custom Drawing coordinates]
    
    AI --> AK([Download output.docx])
    AJ --> AL([Download output.pdf])
    O --> T
    S --> T
```

---

## Core Capabilities

- **Unified Layout Analysis**: Document layout mapping and handwriting transcription are completed in a single multimodal request, avoiding multiple passes and reducing token fees.
- **Dynamic Hierarchy Construction**: The model assigns structural tags (HEADING, SUBHEAD, BODY, BULLET, CENTER, UNDERLN) and calculates typographical margins, font scaling, alignments, and spacings.
- **Refined Text Styling**: Stroke weight, slant, and underline patterns are analyzed at the character run level, applying bold, italic, and underline settings onto the default styles.
- **Smart Alignment for Math**: Multi-line equations and mathematical expressions are centered and aligned around equality operators using calculated padding spaces.
- **Structural Table Recreation**: Extracts borderless and grid tables, dynamically merging cell matrices using rowspan/colspan attributes to handle complex structures.
- **Embedded Sketch Cropping**: Bounding coordinates for sketches, charts, and signatures are cropped out of the original high-resolution inputs and re-embedded inline in their correct reading order.
- **Arrow and Bracket Mapping**: Maps simple hand-drawn arrow vectors and enclosing braces directly into scalable Unicode symbols (Segoe UI Symbol), reducing output file size.
- **Layout Customization and Overrides**: Provides toggle controls between auto-layout detection and manual typography overrides (font size, style, margins, alignment) via the web client.

---

## Multi-Provider Resilience Framework

To ensure high availability under resource limits and API rate constraints, Docify AI implements a multi-provider fallback hierarchy:

1. **Gemini API (Primary)**:
   - Utilizes `gemini-2.5-flash` for high-speed structured JSON generation using schema validation.
   - Implements **API Key Rotation**: Accepts a comma-separated list of keys in `GEMINI_API_KEY` and switches to the next key on 429 quota exhaustion.
   - Implements **Model Failover**: Falls back to `gemini-2.0-flash` on prolonged 503 service unavailability.
   - Retries failed requests up to 4 times with exponential backoff delays (4s, 8s, 16s, 32s).

2. **Groq API (First Fallback)**:
   - Triggered automatically if all Gemini keys/models fail or are unconfigured.
   - Interrogates vision models in sequence:
     - `meta-llama/llama-4-scout-17b-16e-instruct`
     - `meta-llama/llama-4-maverick-17b-128e-instruct`
     - `llama-3.2-90b-vision-preview`
     - `llama-3.2-11b-vision-preview`
   - Dynamically injects Pydantic schema requirements as JSON instructions for models that do not natively support structured outputs.

3. **OpenRouter API (Second Fallback)**:
   - Active if Groq and Gemini fail or are unconfigured.
   - Rotates through free high-capacity vision models:
     - `qwen/qwen2.5-vl-72b-instruct:free`
     - `qwen/qwen2-vl-7b-instruct:free`
     - `meta-llama/llama-3.2-11b-vision-instruct:free`
     - `google/gemini-2.0-flash-exp:free`

4. **Plain OCR and RegEx Fallback**:
   - If structured output calls fail across all APIs, Docify AI requests unstructured, line-by-line transcription.
   - The backend processes the output using regular expressions and prefix mapping (`_parse_text_lines`) to build a clean text layout.

---

## Image Preprocessing and OCR Caching

### Image Preprocessing
Before submission to the vision models, images are processed using standard Pillow transformations:
- **RGB Normalization**: Converts input color spaces (RGBA, CMYK, Palette, Grayscale) into standardized three-channel RGB color values.
- **Lanczos Upscaling**: Small images (width under 1600px) are scaled up using Lanczos filters to make handwritten text clearer for the models.
- **Filter Avoidance**: The engine avoids aggressive sharpening or thresholding to preserve the visual context needed for multimodal comprehension.

### Server-Side OCR Caching
To reduce API overhead and accelerate response times, Docify AI runs a file-hash caching database:
- Computes an MD5 checksum of the processed image file.
- Saves the extracted OCR layout and metadata in JSON format as `uploads/ocr_cache_[MD5_HASH].json`.
- Subsequent conversions of the same document load the parsed structure from the local cache instantly.

---

## Pydantic Layout Schema

Gemini responses are validated against a strict Pydantic model configuration:

```python
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
        description="The formatting tag for text elements: 'HEADING' (major title), 'SUBHEAD' (section label), 'BODY' (normal paragraph), 'BULLET' (list item), 'CENTER' (visually centered line), 'UNDERLN' (underlined text)."
    )
    bold: Optional[bool] = Field(
        None,
        description="True if the text is bold or written significantly darker than normal."
    )
    italic: Optional[bool] = Field(
        None,
        description="True if the text is italicized."
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
        description="Estimated left indentation of the text element in centimeters."
    )
    table_data: Optional[List[List[str]]] = Field(
        None,
        description="A list of rows representing the cells of the table."
    )
    borderless: Optional[bool] = Field(
        None,
        description="True if the table is a layout grid and should be rendered without visible borders."
    )
    bbox: Optional[List[float]] = Field(
        None,
        description="Normalized bounding box [x1, y1, x2, y2] (0.0 to 1.0) of the drawing in the image."
    )
    description: Optional[str] = Field(
        None,
        description="A brief description of the drawing or diagram."
    )
    is_simple_arrow: Optional[bool] = Field(
        None,
        description="True if the drawing is just a single plain hand-drawn arrow."
    )
    arrow_direction: Optional[str] = Field(
        None,
        description="Direction the arrow points: 'right', 'left', 'up', 'down', 'up-right', etc."
    )
    is_simple_bracket: Optional[bool] = Field(
        None,
        description="True if the drawing is just a single grouping bracket or brace."
    )
    bracket_style: Optional[str] = Field(
        None,
        description="The style of bracket: 'curly', 'square', 'plain'."
    )
    bracket_side: Optional[str] = Field(
        None,
        description="Which side the bracket appears relative to the text: 'left', 'right'."
    )

class DocumentLayout(BaseModel):
    page_margin_cm: float = Field(
        default=2.54,
        description="Estimated page margin in centimeters."
    )
    line_spacing: float = Field(
        default=1.15,
        description="Line spacing multiplier: 1.0, 1.15, 1.5, 2.0."
    )
    elements: List[DocElement] = Field(
        ...,
        description="The ordered list of all document elements from top to bottom."
    )
```

---

## Element Mapping and Output Assembly

The table below outlines the default formatting rules used by the mapping controller when translating layout tags into document outputs:

| Layout Tag | Base Weight | Slant | Font Size | Paragraph Alignment | Left Indent | Space Before | Space After |
| :--- | :---: | :---: | :---: | :--- | :---: | :---: | :---: |
| `HEADING` | Bold | Normal | 16 pt | Left | 0.0 cm | 12 pt | 6 pt |
| `SUBHEAD` | Bold | Normal | 13 pt | Left | 0.0 cm | 8 pt | 4 pt |
| `BODY` | Normal | Normal | 12 pt | Left | 0.0 cm | 0 pt | 3 pt |
| `BULLET` | Normal | Normal | 12 pt | Left | 0.5 cm | 0 pt | 3 pt |
| `CENTER` | Normal | Normal | 12 pt | Center | 0.0 cm | 4 pt | 4 pt |
| `UNDERLN` | Normal | Normal | 12 pt | Left | 0.0 cm | 0 pt | 3 pt |
| `TABLE` | Normal | Normal | 11 pt | Left | 0.0 cm | 4 pt | 4 pt |

### Custom Processing Logic
- **Signature Blocks**: Uses borderless tables with empty headers to place side-by-side signature components next to each other.
- **Dynamic Text Shrinkage**: Footer notes and metadata blocks are scaled down to 8.5 pt and 9.5 pt dynamically to match professional layouts.
- **Double Space Cleanup**: Cleans up multiple spaces and converts them to non-breaking spaces (`\xA0`) to keep formatting aligned.

---

## Technology Stack

| Layer | Dependency | Version | Purpose |
| :--- | :--- | :--- | :--- |
| **API & Routing** | FastAPI | 0.110.0+ | Server framework and REST API endpoints |
| **ASGI Engine** | Uvicorn | 0.28.0+ | Server runner and hot-reloader |
| **Vision AI** | google-genai | 0.1.0+ | Gemini models SDK connection |
| **Failover Models** | groq / openai | 0.5.0+ / 1.12.0+ | Fallback integration for secondary engines |
| **PDF Extraction** | pypdfium2 | 4.28.0+ | Multi-page PDF layout splitter |
| **Image Processing** | Pillow (PIL) | 10.2.0+ | Image sizing, formats, and canvas rendering |
| **Word Generator** | python-docx | 1.1.0+ | Native XML Word layout builder |
| **PDF Generator** | ReportLab | 4.1.0+ | Custom coordinates vector canvas generator |
| **Data Validation** | Pydantic v2 | 2.6.0+ | JSON layout validation |
| **Configuration** | python-dotenv | 1.0.0+ | Environment variables loader |
| **Frontend UI** | HTML5 / CSS3 / JS | Native | Modern dark/light responsive interface |

---

## Directory Structure

```
docify/
├── main.py                  # Core backend app: handles routing, vision, document generation
├── run.bat                  # Windows runner script
├── requirements.txt         # Project package dependencies
├── .env                     # Configuration keys (excluded from source control)
├── .gitignore               # Ignored files configuration
├── static/                  # Shared client assets
│   ├── css/
│   │   ├── styles.css       # Main application styling (dark/light mode variables)
│   │   └── landing.css      # Landing presentation style sheet
│   └── js/
│       ├── main.js          # Core interface logic: upload flow and parameter mapping
│       ├── landing.js       # Animations and navigation for landing page
│       ├── about.js         # Technical informational display operations
│       ├── history.js       # client-side conversion history management
│       ├── login.js         # Authorization layout functions
│       └── theme.js         # Dynamic dark/light color scheme toggle controller
├── templates/               # Server-side HTML render layouts
│   ├── landing.html         # Front landing portal
│   ├── index.html           # Core converter system page
│   ├── about.html           # Developer about info page
│   ├── history.html         # Saved conversion record logs page
│   └── login.html           # Authentication portal page
├── uploads/                 # Storage for source files and cropped sketches (ignored)
└── outputs/                 # Storage for generated word/pdf download packages (ignored)
```

---

## Installation Guide

### Prerequisites
- Python 3.9 or higher.
- `pip` package manager.
- **Poppler** utility binaries (required by `pdf2image` to extract pages from PDF files).
- An API Key from [Google AI Studio](https://aistudio.google.com/).
- Optional: Groq or OpenRouter API credentials for failover configurations.

---

### Installing Poppler

#### Windows
1. Download the latest pre-compiled binary package from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/).
2. Extract the directory contents to an accessible path (e.g., `C:\poppler`).
3. Add the path to the extracted `bin` folder (e.g., `C:\poppler\Library\bin` or `C:\poppler\bin`) to your system **Path** environment variable:
   - Search for **Environment Variables** in the Windows taskbar.
   - Under System Variables, click **Path** and select Edit.
   - Click New and enter the path to the Poppler `bin` directory.
   - Click OK to save the changes.
4. Restart your terminal application to apply the new path settings.

#### macOS
Install poppler using the Homebrew package manager:
```bash
brew install poppler
```

#### Linux
Install poppler using your distribution's package manager:
```bash
# Ubuntu / Debian
sudo apt-get update
sudo apt-get install -y poppler-utils

# Fedora / CentOS
sudo dnf install poppler-utils
```

---

### Step-by-Step Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/meetchauhan17/docify.git
   cd docify
   ```

2. **Establish a Python virtual environment**:
   ```bash
   python -m venv venv
   ```

3. **Activate the environment**:
   - **Windows (PowerShell)**:
     ```powershell
     .\venv\Scripts\Activate.ps1
     ```
   - **Windows (CMD)**:
     ```cmd
     .\venv\Scripts\activate.bat
     ```
   - **macOS / Linux**:
     ```bash
     source venv/bin/activate
     ```

4. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

5. **Configure environment settings**:
   Create a `.env` file in the project root:
   ```env
   # API Keys configuration (comma-separated for key rotation on Gemini)
   GEMINI_API_KEY=key_one,key_two
   
   # Optional Fallback API Keys
   GROQ_API_KEY=your_groq_api_key_here
   OPENROUTER_API_KEY=your_openrouter_api_key_here
   ```

6. **Initialize folders**:
   ```bash
   mkdir uploads
   mkdir outputs
   ```

---

## Running the Application

### Option A: Windows Launcher
Run the setup batch file:
```cmd
run.bat
```

### Option B: Manual CLI Run
Start the Uvicorn ASGI server:
```bash
uvicorn main:app --reload --port 8000
```
Open a web browser and navigate to `http://localhost:8000`.

---

## Web Interface Reference

Docify AI exposes a multi-page web application served by FastAPI with Jinja2 templating:

| Route Path | View Interface | Technical Role |
| :--- | :--- | :--- |
| `/` | Landing page | Project overview, interactive features overview, and system portal access. |
| `/convert` | Converter page | Main interface for file upload, format selector, auto/manual toggles, and processing triggers. |
| `/history` | History page | Lists previous conversion records loaded client-side via browser `localStorage`. |
| `/about` | Technical Information | Full technical walkthrough outlining elements parsing rules and configuration info. |
| `/login` | Authentication page | Login portal structure for user verification systems. |

---

## API Reference

### POST /convert

Submits a document file to be parsed and downloads the formatted output.

- **Content-Type**: `multipart/form-data`

#### Request Parameters

| Parameter | Data Type | Required | Default Value | Functional Role |
| :--- | :--- | :---: | :--- | :--- |
| `file` | Binary | Yes | - | Source file upload (`PNG`, `JPG`, `JPEG`, `WEBP`, or `PDF`). |
| `outtype` | String | No | `docx` | Target export format: `docx` or `pdf`. |
| `mode` | String | No | `preserve` | Text structure layout: `preserve` respects original lines; `flow` combines paragraphs. |
| `auto_format` | String | No | `true` | When `true`, uses AI layout extraction. When `false`, uses the manual overrides below. |
| `font_family` | String | No | `Calibri` | Manual override font family (e.g. `Arial`, `Times New Roman`, `Georgia`). |
| `font_size` | Float | No | `12.0` | Manual override font size in points. |
| `line_spacing` | Float | No | `1.15` | Manual override line spacing multiplier: `1.0`, `1.15`, `1.5`, `2.0`. |
| `para_spacing` | Float | No | `8.0` | Manual override space after paragraphs in points. |
| `first_line_indent` | Float | No | `0.0` | Manual override first line indent in centimeters. |
| `page_margin` | Float | No | `2.54` | Manual override page margin in centimeters. |
| `text_align` | String | No | `left` | Manual override text alignment: `left`, `center`, `right`, `justify`. |
| `text_bold` | String | No | `false` | Manual override: forces all text to bold (`true`/`false`). |
| `text_italic` | String | No | `false` | Manual override: forces all text to italic (`true`/`false`). |
| `text_underline` | String | No | `false` | Manual override: forces all text to underline (`true`/`false`). |

#### Example Client Submissions

##### cURL command:
```bash
curl -X POST http://localhost:8000/convert \
  -F "file=@handwritten_sheet.jpg" \
  -F "outtype=docx" \
  -F "auto_format=true" \
  --output document_result.docx
```

##### Python script:
```python
import requests

url = "http://localhost:8000/convert"
files = {"file": open("handwritten_sheet.jpg", "rb")}
data = {
    "outtype": "pdf",
    "auto_format": "true",
    "mode": "preserve"
}

response = requests.post(url, files=files, data=data)

if response.status_code == 200:
    with open("document_result.pdf", "wb") as f:
        f.write(response.content)
    print("Document compiled successfully.")
else:
    print(f"Error: {response.json()}")
```

---

## Configuration Reference

Docify AI loads its primary runtime values from the system environment or your local `.env` file:

| Environment Variable | Required | Description |
| :--- | :---: | :--- |
| `GEMINI_API_KEY` | Yes | Comma-separated list of Google Gemini API keys. |
| `GROQ_API_KEY` | No | API Key for backup Llama vision models from Groq. |
| `OPENROUTER_API_KEY` | No | API Key for backup free vision models from OpenRouter. |

### Internal System Constants
These variables can be adjusted directly in `main.py` if needed:

| Constant | Default Value | Description |
| :--- | :--- | :--- |
| `_PRIMARY_MODEL` | `gemini-2.5-flash` | The primary vision model for processing OCR structured layout JSON. |
| `_FALLBACK_MODEL` | `gemini-2.0-flash` | Backup Gemini model tried when the primary returns a 503 error. |
| `UPLOAD_DIR` | `"uploads"` | Folder used to store uploaded files and cropped drawing assets. |
| `OUTPUT_DIR` | `"outputs"` | Folder where generated Word and PDF files are saved. |
| `TIMEOUT` | `120,000 ms` | Maximum time allowed per vision API call before timeout. |

---

## License

This project is licensed under the terms of the MIT License. See the [LICENSE](LICENSE) file for complete details.
