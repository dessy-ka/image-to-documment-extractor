# Adaptive Batch Image-to-Document Extractor

A stateless, in-memory web application that converts **1–15 document images** into editable, styled **Excel (`.xlsx`)** workbooks or **Word (`.docx`)** documents using **Google Gemini Vision** OCR and **FastAPI**.

---

## Features

- **Batch Processing (Up to 15 Images):** Queue, reorder, or remove up to 15 images via drag-and-drop, file picker, or mobile camera capture.
- **Client-Side Image Optimization:** Automatically resizes images exceeding `1600px` on the longest edge and compresses them to `85%` JPEG quality before upload.
- **Adaptive Layout Detection:**
  - `Auto-Detect`: Automatically chooses Excel or Word based on majority vote across valid pages (defaults to Excel on ties).
  - `Force Excel (.xlsx)`: Extracts tabular data, merges matching-schema pages into `Extracted_Records` (deduplicating repeated headers), and places distinct schemas on separate `Page_<N>` sheets.
  - `Force Word (.docx)`: Preserves prose paragraphs, alignment, bold formatting, borderless `split_row` layouts, and embedded tables.
- **Side-by-Side Review & Interactive Editing:** Inspect extracted content alongside the original reference image, edit cells/headers/rows/columns (Excel) or blocks/alignment/tables (Word), and replace or re-extract failed pages in place.
- **Stateless & In-Memory:** All image validation (`Pillow`), EXIF orientation normalization, OCR, and `.xlsx`/`.docx` generation happen in memory with no database or disk persistence.

---

## Tech Stack

- **Backend:** Python 3.11+, FastAPI, Uvicorn, Pydantic, Pillow, `google-genai`, `openpyxl`, `python-docx`, `python-dotenv`
- **Frontend:** Single-file HTML5, Vanilla JavaScript, and local Tailwind CSS (`static/index.html`, `static/tailwind.js`)
- **AI / OCR:** Google Gemini (`gemini-3.8-flash` by default, overridable via `GEMINI_MODEL`)

---

## Quick Start

### Prerequisites

- **Python 3.11+**
- A **Google Gemini API key**

### 1. Configure Environment Variables

Copy `.env.example` to `.env` in the project root and add your API key:

```env
GEMINI_API_KEY=your_gemini_api_key_here
```

### 2. Run the Application

#### Option A: Startup Scripts (Recommended)

- **Windows (PowerShell):**
  ```powershell
  .\app_script\windows\start_server.ps1
  ```
  *(Or double-click `app_script\windows\start_server.bat`.)*
- **Linux / macOS:**
  ```bash
  bash app_script/linux_mac/install_requirements.sh
  bash app_script/linux_mac/start_server.sh
  ```

The startup scripts automatically create `.venv`, install `requirements.txt`, copy `.env.example` to `.env` if missing, open `http://127.0.0.1:8000` in your browser, and start the FastAPI server.

#### Option B: Manual Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn main:app --reload --host 127.0.0.1 --port 8000
```

Then open `http://127.0.0.1:8000` in your browser.

---

## Project Structure

```text
├── main.py                      # FastAPI backend, Gemini OCR, merging, & .xlsx/.docx builders
├── requirements.txt             # Python dependencies
├── .env.example                 # Environment variable template (GEMINI_API_KEY)
├── static/
│   ├── index.html               # Single-file responsive web UI
│   └── tailwind.js              # Local Tailwind CSS browser runtime
├── app_script/
│   ├── windows/                 # Windows install & startup scripts (.ps1, .bat)
│   └── linux_mac/               # Linux/macOS install & startup scripts (.sh)
└── docs/
    ├── SETUP.md                 # Detailed setup & troubleshooting guide
    └── IMPLEMENTATION.md        # Technical specification & architecture
```

---

## API Endpoints

| Method | Path | Description |
| :--- | :--- | :--- |
| `GET` | `/` | Serves the web UI (`static/index.html`) |
| `POST` | `/extract` | Accepts `files` (1–15 images, max 10 MB each) and `mode` (`auto`, `excel`, `word`); returns `MergedDocumentPayload` JSON |
| `POST` | `/export` | Accepts `MergedDocumentPayload` JSON and streams a generated `.xlsx` or `.docx` file |

---

## Documentation

- [Setup & Troubleshooting Guide](docs/SETUP.md)
- [Implementation & Technical Specification](docs/IMPLEMENTATION.md)
