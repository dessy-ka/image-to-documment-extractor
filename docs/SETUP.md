# Setup Guide — Adaptive Batch Image-to-Document Extractor

## Prerequisites
- **Python 3.11+** (verified on Python 3.14)
- A **Google Gemini API key** with access to `gemini-3.8-flash`

---

## 1. Configure Environment Variables
1. Copy `.env.example` to `.env` in the project root:
   ```powershell
   Copy-Item .env.example .env
   ```
2. Open `.env` and set your Gemini API key:
   ```env
   GEMINI_API_KEY=your_gemini_api_key_here
   ```

---

## 2. Start the Application

### Option A: Automatic Startup Script (Recommended)
**Windows:** Run the PowerShell script from the project root, or double-click `app_script\windows\start_server.bat` in File Explorer:
```powershell
.\app_script\windows\start_server.ps1
```
You can also install or update dependencies standalone using `app_script\windows\install_requirements.ps1` (or `app_script\windows\install_requirements.bat`).

**Linux / macOS:**
```bash
bash app_script/linux_mac/install_requirements.sh
bash app_script/linux_mac/start_server.sh
```

The startup script automatically:
- Creates `.venv` and installs `requirements.txt` if `.venv` does not exist.
- Creates `.env` from `.env.example` if `.env` is missing.
- Opens `http://127.0.0.1:8000` in your default browser and starts `uvicorn main:app --reload`.

### Option B: Manual Setup & Startup
```powershell
# 1. Create and activate a virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the FastAPI server
uvicorn main:app --reload --host 127.0.0.1 --port 8000
```
Then open `http://127.0.0.1:8000` in your browser.

---

## 3. Usage Workflow
1. **Queue Images (Stage 1):** Drag and drop up to 15 images, select files, or tap **Add Camera Photo** on mobile. Reorder or remove pages, choose `Auto-Detect`, `Force Excel (.xlsx)`, or `Force Word (.docx)`, and click **Extract & Preview**.
2. **Review & Edit (Stage 2):** Compare extracted content side-by-side with the original reference image. Edit cells, add/delete rows or columns (Excel), or adjust alignment, bold formatting, split rows, and embedded tables (Word). If any page fails OCR, dismiss the warning or replace and re-extract that page in place.
3. **Export (Stage 3):** Edit the output filename if desired and click **Download Excel (.xlsx)** or **Download Word (.docx)**.

---

## 4. Troubleshooting
- **HTTP 500 (`GEMINI_API_KEY environment variable is not configured`):** Ensure `.env` exists in the project root and contains a valid `GEMINI_API_KEY`.
- **HTTP 413 (File too large):** Each uploaded image must be 10 MB or smaller (the browser UI automatically downscales images to 1600px max edge at 85% JPEG quality before upload).
- **HTTP 429 (Rate limit exceeded):** The server limits extraction requests to 15 requests per 60 seconds per IP. Wait a minute and retry.
- **PowerShell execution policy error:** Run `app_script\windows\start_server.bat` instead, or run `powershell -ExecutionPolicy Bypass -File .\app_script\windows\start_server.ps1`.
