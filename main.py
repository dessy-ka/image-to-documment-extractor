"""Adaptive Batch Image-to-Document Extractor backend."""

import asyncio
from collections import defaultdict
import io
import logging
import os
import re
import time
from typing import Any, Literal, Optional

from dotenv import load_dotenv
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps, UnidentifiedImageError
from pydantic import BaseModel, Field

load_dotenv(override=True)
logger = logging.getLogger("uvicorn.error")

MAX_BATCH_SIZE = 15
MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024  # 10 MB per image
RATE_LIMIT_REQUESTS = 15
RATE_LIMIT_WINDOW_SECONDS = 60
GEMINI_CONCURRENCY = 5
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")

# Guard against decompression bombs in Pillow
Image.MAX_IMAGE_PIXELS = 40_000_000

app = FastAPI(title="Adaptive Batch Image-to-Document Extractor")
app.mount(
    "/static",
    StaticFiles(directory=os.path.join(os.path.dirname(__file__), "static")),
    name="static",
)

_rate_limit_store: dict[str, list[float]] = defaultdict(list)

CURRENCY_PREFIX_RE = re.compile(
    r"^(?:GHS|₵|\$|€|£)\s*", flags=re.IGNORECASE
)
INVALID_SHEET_CHARS_RE = re.compile(r"[:\\/?*\[\]]")
SLUG_UNSAFE_RE = re.compile(r"[^a-zA-Z0-9_-]+")


# ---------------------------------------------------------------------------
# Pydantic Schemas
# ---------------------------------------------------------------------------


class ProseBlock(BaseModel):
    """Represents a single logical block in a prose/letter/mixed document."""

    kind: Literal["single", "split_row", "table"]
    text: str = ""
    alignment: Literal["left", "center", "right"] = "left"
    is_bold: bool = False
    left_text: str = ""
    right_text: str = ""
    table_headers: list[str] = Field(default_factory=list)
    table_rows: list[list[str]] = Field(default_factory=list)


class PageExtractionResult(BaseModel):
    """Structured per-page output schema returned by Gemini Vision OCR."""

    layout_type: Literal["tabular", "prose"]
    suggested_filename: str = "extracted_document"
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    blocks: list[ProseBlock] = Field(default_factory=list)


class ExcelSheetPayload(BaseModel):
    """Represents one worksheet in an Excel export."""

    sheet_name: str
    headers: list[str]
    rows: list[list[str]]


class PageError(BaseModel):
    """Describes a per-page extraction or validation failure."""

    page_index: int
    filename: str
    reason: str


class MergedDocumentPayload(BaseModel):
    """Merged document payload returned by /extract and accepted by /export."""

    output_format: Literal["excel", "word"]
    filename: str
    pages_processed: int
    failed_pages: list[PageError] = Field(default_factory=list)
    sheets: Optional[list[ExcelSheetPayload]] = None
    pages: Optional[list[list[ProseBlock]]] = None


# ---------------------------------------------------------------------------
# Helpers: Rate Limiting, Sanitization, & Image Pre-Processing
# ---------------------------------------------------------------------------


def _get_client_ip(request: Request) -> str:
    forwarded = request.headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.client.host if request.client else "unknown"


def _check_rate_limit(request: Request) -> None:
    ip = _get_client_ip(request)
    now = time.monotonic()
    cutoff = now - RATE_LIMIT_WINDOW_SECONDS
    recent = [ts for ts in _rate_limit_store[ip] if ts > cutoff]
    if len(recent) >= RATE_LIMIT_REQUESTS:
        _rate_limit_store[ip] = recent
        raise HTTPException(
            status_code=429,
            detail=(
                f"Rate limit exceeded ({RATE_LIMIT_REQUESTS} requests per "
                f"{RATE_LIMIT_WINDOW_SECONDS}s). Please wait and try again."
            ),
        )
    recent.append(now)
    _rate_limit_store[ip] = recent


def _sanitize_filename(name: str, fallback: str = "extracted_document") -> str:
    cleaned = SLUG_UNSAFE_RE.sub("_", name.strip()).strip("_-")
    return cleaned[:80] if cleaned else fallback


def _validate_and_normalize_image(raw_bytes: bytes) -> bytes:
    """Verify image magic bytes via Pillow, fix EXIF orientation, and return JPEG bytes."""
    if not raw_bytes:
        raise ValueError("Uploaded file is empty.")

    try:
        with Image.open(io.BytesIO(raw_bytes)) as img_check:
            img_check.verify()
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise ValueError("File is not a valid or supported image.") from exc

    try:
        with Image.open(io.BytesIO(raw_bytes)) as img:
            transposed = ImageOps.exif_transpose(img)
            if transposed.mode in ("RGBA", "LA", "P"):
                rgba = transposed.convert("RGBA")
                background = Image.new("RGB", rgba.size, (255, 255, 255))
                background.paste(rgba, mask=rgba.split()[3])
                rgb = background
            elif transposed.mode != "RGB":
                rgb = transposed.convert("RGB")
            else:
                rgb = transposed

            out_buf = io.BytesIO()
            rgb.save(out_buf, format="JPEG", quality=90)
            return out_buf.getvalue()
    except (OSError, ValueError) as exc:
        raise ValueError(f"Failed to decode or normalize image: {exc}") from exc


# ---------------------------------------------------------------------------
# Helpers: Gemini Vision Extraction
# ---------------------------------------------------------------------------


def _build_extraction_prompt(mode: Literal["auto", "excel", "word"]) -> str:
    base = (
        "You are an expert document OCR and layout-extraction engine. "
        "Transcribe ALL visible text accurately—including handwriting, faded ink, "
        "and low-light or slightly skewed photos—without inventing or skipping content. "
        "Provide a short, clean snake_case `suggested_filename` (no extension) describing the document.\n\n"
    )
    if mode == "excel":
        return base + (
            "MODE: EXCEL (TABULAR). Set `layout_type` to \"tabular\". "
            "Extract all records/data into `headers` (list of column names) and `rows` "
            "(list of rows, each row having the exact same number of cells as `headers`). "
            "Leave `blocks` empty."
        )
    if mode == "word":
        return base + (
            "MODE: WORD (PROSE). Set `layout_type` to \"prose\". "
            "Extract the document into `blocks` in top-to-bottom reading order:\n"
            "- Use kind=\"single\" for standard paragraphs, centered titles/letterheads, "
            "right-aligned dates/addresses/signatures, with accurate `alignment` (\"left\", \"center\", \"right\") "
            "and `is_bold`.\n"
            "- Use kind=\"split_row\" when a single horizontal line has distinct text on the far left "
            "and far right (e.g., reference number on left, date on right), populating `left_text` and `right_text`.\n"
            "- Use kind=\"table\" for any embedded table, populating `table_headers` and `table_rows`.\n"
            "Leave top-level `headers` and `rows` empty."
        )
    return base + (
        "MODE: AUTO-DETECT.\n"
        "- If the image is primarily a table, ledger, receipt, invoice line-item sheet, inventory list, "
        "or attendance log, set `layout_type` to \"tabular\" and populate `headers` and `rows` "
        "(every row matching `headers` length). Leave `blocks` empty.\n"
        "- If the image is a letter, memo, notice, report, or prose document (even if it contains an embedded table), "
        "set `layout_type` to \"prose\" and populate `blocks`:\n"
        "  * kind=\"single\": `text`, `alignment` (\"left\" | \"center\" | \"right\"), `is_bold`.\n"
        "  * kind=\"split_row\": `left_text`, `right_text` when one horizontal line has distinct left and right text.\n"
        "  * kind=\"table\": `table_headers`, `table_rows` for embedded tables.\n"
        "  Leave top-level `headers` and `rows` empty."
    )


def _format_gemini_error(exc: Exception) -> str:
    """Extract a concise, user-friendly error message from Gemini SDK exceptions."""
    msg = getattr(exc, "message", None) or str(exc)
    code = getattr(exc, "code", None)
    upper = msg.upper()
    if code in (401, 403) or "API_KEY" in upper or "PERMISSION_DENIED" in upper:
        return "Invalid or unauthorized Gemini API key. Check GEMINI_API_KEY in .env."
    if code == 429 or "RESOURCE_EXHAUSTED" in upper or "QUOTA" in upper:
        return "Gemini API rate limit or quota exceeded. Please wait a moment and retry."
    if code == 404 or "NOT_FOUND" in upper:
        return f"Gemini model '{GEMINI_MODEL}' is unavailable: {msg}"
    return msg[:240] if len(msg) > 240 else msg


async def _extract_single_page(
    client: genai.Client,
    semaphore: asyncio.Semaphore,
    page_index: int,
    filename: str,
    image_bytes: bytes,
    mode: Literal["auto", "excel", "word"],
) -> PageExtractionResult:
    prompt = _build_extraction_prompt(mode)
    async with semaphore:
        response = await client.aio.models.generate_content(
            model=GEMINI_MODEL,
            contents=[
                types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg"),
                prompt,
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=PageExtractionResult,
                temperature=0.1,
            ),
        )

    if response.parsed is not None and isinstance(response.parsed, PageExtractionResult):
        return response.parsed
    if response.text:
        return PageExtractionResult.model_validate_json(response.text)
    raise ValueError(f"Empty response from Gemini for page {page_index + 1} ({filename}).")


# ---------------------------------------------------------------------------
# Helpers: Batch Merging
# ---------------------------------------------------------------------------


def _normalize_row(row: list[str], width: int) -> list[str]:
    if width <= 0:
        return [str(c) for c in row]
    padded = [str(c) for c in row[:width]]
    while len(padded) < width:
        padded.append("")
    return padded


def _headers_match(h1: list[str], h2: list[str]) -> bool:
    return [c.strip().lower() for c in h1] == [c.strip().lower() for c in h2]


def _prose_to_sheets(
    blocks: list[ProseBlock], page_num: int
) -> list[ExcelSheetPayload]:
    """Convert a prose page into one or more Excel sheets without losing content."""
    sheets: list[ExcelSheetPayload] = []
    text_rows: list[list[str]] = []

    for block in blocks:
        if block.kind == "table" and (block.table_headers or block.table_rows):
            headers = block.table_headers or [
                f"Col_{i + 1}"
                for i in range(max((len(r) for r in block.table_rows), default=1))
            ]
            rows = [_normalize_row(r, len(headers)) for r in block.table_rows]
            suffix = f"_{len(sheets) + 1}" if sheets else ""
            sheets.append(
                ExcelSheetPayload(
                    sheet_name=f"Page_{page_num}{suffix}",
                    headers=headers,
                    rows=rows,
                )
            )
        elif block.kind == "split_row":
            if block.left_text or block.right_text:
                text_rows.append([block.left_text, block.right_text])
        elif block.kind == "single" and block.text.strip():
            text_rows.append([block.text.strip(), ""])

    if text_rows:
        sheets.append(
            ExcelSheetPayload(
                sheet_name=f"Page_{page_num}_Text",
                headers=["Content", "Secondary"],
                rows=text_rows,
            )
        )
    return sheets


def _merge_excel_pages(
    valid_pages: list[tuple[int, PageExtractionResult]],
) -> list[ExcelSheetPayload]:
    sheets: list[ExcelSheetPayload] = []
    primary_headers: Optional[list[str]] = None

    for idx, page in valid_pages:
        if page.layout_type == "tabular" and (page.headers or page.rows):
            headers = page.headers or [
                f"Column_{i + 1}"
                for i in range(max((len(r) for r in page.rows), default=1))
            ]
            width = len(headers)
            rows = [_normalize_row(r, width) for r in page.rows]

            if primary_headers is None:
                primary_headers = headers
                sheets.append(
                    ExcelSheetPayload(
                        sheet_name="Extracted_Records",
                        headers=headers,
                        rows=rows,
                    )
                )
            elif _headers_match(primary_headers, headers):
                if rows and _headers_match(primary_headers, rows[0]):
                    rows = rows[1:]
                sheets[0].rows.extend(rows)
            else:
                sheets.append(
                    ExcelSheetPayload(
                        sheet_name=f"Page_{idx + 1}",
                        headers=headers,
                        rows=rows,
                    )
                )
        else:
            converted = _prose_to_sheets(page.blocks, idx + 1)
            for sheet in converted:
                if (
                    primary_headers is not None
                    and _headers_match(primary_headers, sheet.headers)
                ):
                    rows = sheet.rows
                    if rows and _headers_match(primary_headers, rows[0]):
                        rows = rows[1:]
                    sheets[0].rows.extend(rows)
                else:
                    if primary_headers is None:
                        primary_headers = sheet.headers
                        sheet.sheet_name = "Extracted_Records"
                    sheets.append(sheet)

    if not sheets:
        sheets.append(
            ExcelSheetPayload(
                sheet_name="Extracted_Records",
                headers=["Column_1"],
                rows=[],
            )
        )
    return sheets


def _merge_word_pages(
    valid_pages: list[tuple[int, PageExtractionResult]],
) -> list[list[ProseBlock]]:
    pages: list[list[ProseBlock]] = []
    for _, page in valid_pages:
        if page.layout_type == "prose" and page.blocks:
            pages.append(page.blocks)
        elif page.headers or page.rows:
            pages.append(
                [
                    ProseBlock(
                        kind="table",
                        table_headers=page.headers,
                        table_rows=page.rows,
                    )
                ]
            )
        else:
            pages.append([ProseBlock(kind="single", text="")])
    return pages


# ---------------------------------------------------------------------------
# Helpers: Excel & Word File Generators (Strictly In-Memory)
# ---------------------------------------------------------------------------


def _sanitize_sheet_name(name: str, existing: set[str]) -> str:
    cleaned = INVALID_SHEET_CHARS_RE.sub("_", name.strip()) or "Sheet"
    base = cleaned[:31]
    candidate = base
    counter = 2
    while candidate in existing:
        suffix = f"_{counter}"
        candidate = f"{base[: 31 - len(suffix)]}{suffix}"
        counter += 1
    existing.add(candidate)
    return candidate


def _parse_numeric_cell(raw: str) -> tuple[Any, Optional[str]]:
    """Parse numeric/currency strings into int/float while preserving leading-zero IDs."""
    text = raw.strip()
    if not text:
        return "", None

    has_currency = bool(CURRENCY_PREFIX_RE.match(text))
    cleaned = CURRENCY_PREFIX_RE.sub("", text).strip()

    is_negative = False
    if cleaned.startswith("(") and cleaned.endswith(")"):
        is_negative = True
        cleaned = cleaned[1:-1].strip()
    elif cleaned.startswith("-"):
        is_negative = True
        cleaned = cleaned[1:].strip()

    # Re-check currency inside parentheses e.g. ($1,200.00)
    if CURRENCY_PREFIX_RE.match(cleaned):
        has_currency = True
        cleaned = CURRENCY_PREFIX_RE.sub("", cleaned).strip()

    cleaned = cleaned.replace(",", "").replace(" ", "")
    if not cleaned:
        return raw, None

    # Preserve leading-zero identifiers (e.g. "00412", "0244123456") as text
    if (
        not has_currency
        and not is_negative
        and "." not in cleaned
        and cleaned.isdigit()
        and len(cleaned) > 1
        and cleaned.startswith("0")
    ):
        return raw, None

    if re.fullmatch(r"\d+", cleaned):
        val = int(cleaned)
        if is_negative:
            val = -val
        return (float(val), "#,##0.00") if has_currency else (val, "#,##0")

    if re.fullmatch(r"\d*\.\d+", cleaned):
        val_f = float(cleaned)
        if is_negative:
            val_f = -val_f
        return val_f, "#,##0.00"

    return raw, None


def build_excel(payload: MergedDocumentPayload) -> io.BytesIO:
    """Build an in-memory .xlsx workbook from MergedDocumentPayload."""
    if not payload.sheets:
        raise HTTPException(status_code=400, detail="No sheets provided for Excel export.")

    wb = Workbook()
    # Remove default sheet after creating sheets
    default_ws = wb.active
    used_names: set[str] = set()

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell_font = Font(name="Calibri", size=11)
    cell_align = Alignment(vertical="center", wrap_text=True)

    for sheet_data in payload.sheets:
        title = _sanitize_sheet_name(sheet_data.sheet_name, used_names)
        ws = wb.create_sheet(title=title)

        headers = sheet_data.headers or []
        col_widths: dict[int, int] = {}

        if headers:
            ws.append(headers)
            for col_idx, h_val in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                col_widths[col_idx] = max(col_widths.get(col_idx, 10), len(str(h_val)))
            ws.row_dimensions[1].height = 24

        start_row = 2 if headers else 1
        for r_offset, row_vals in enumerate(sheet_data.rows):
            row_idx = start_row + r_offset
            for col_idx, raw_val in enumerate(row_vals, start=1):
                parsed_val, num_fmt = _parse_numeric_cell(str(raw_val))
                cell = ws.cell(row=row_idx, column=col_idx, value=parsed_val)
                cell.font = cell_font
                cell.alignment = cell_align
                if num_fmt:
                    cell.number_format = num_fmt
                display_len = len(str(raw_val))
                col_widths[col_idx] = max(col_widths.get(col_idx, 10), display_len)

        for col_idx, max_len in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 45)

    if default_ws is not None and len(wb.worksheets) > 1:
        wb.remove(default_ws)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _strip_table_borders(table: Any) -> None:
    """Remove all borders from a python-docx table at the XML level."""
    tbl_pr = table._tbl.tblPr
    existing = tbl_pr.find(qn("w:tblBorders"))
    if existing is not None:
        tbl_pr.remove(existing)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        "</w:tblBorders>"
    )
    tbl_pr.append(borders)


def _set_cell_padding(cell: Any, top: int = 80, bottom: int = 80, left: int = 120, right: int = 120) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def _shade_cell(cell: Any, hex_color: str) -> None:
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def build_word(payload: MergedDocumentPayload) -> io.BytesIO:
    """Build an in-memory .docx document from MergedDocumentPayload."""
    if not payload.pages:
        raise HTTPException(status_code=400, detail="No pages provided for Word export.")

    doc = Document()
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }

    for page_idx, blocks in enumerate(payload.pages):
        if page_idx > 0:
            doc.add_page_break()

        for block in blocks:
            if block.kind == "single":
                p = doc.add_paragraph()
                p.alignment = align_map.get(block.alignment, WD_ALIGN_PARAGRAPH.LEFT)
                p.paragraph_format.space_after = Pt(6)
                run = p.add_run(block.text)
                run.bold = block.is_bold

            elif block.kind == "split_row":
                tbl = doc.add_table(rows=1, cols=2)
                tbl.autofit = False
                _strip_table_borders(tbl)
                left_cell = tbl.cell(0, 0)
                right_cell = tbl.cell(0, 1)
                left_cell.width = Inches(3.25)
                right_cell.width = Inches(3.25)

                lp = left_cell.paragraphs[0]
                lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                lp.paragraph_format.space_after = Pt(4)
                lp.add_run(block.left_text)

                rp = right_cell.paragraphs[0]
                rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                rp.paragraph_format.space_after = Pt(4)
                rp.add_run(block.right_text)

            elif block.kind == "table":
                headers = block.table_headers or []
                rows = block.table_rows or []
                num_cols = max(
                    len(headers),
                    max((len(r) for r in rows), default=0),
                    1,
                )
                total_rows = (1 if headers else 0) + len(rows)
                if total_rows == 0:
                    continue

                tbl = doc.add_table(rows=total_rows, cols=num_cols)
                tbl.style = "Table Grid"

                row_cursor = 0
                if headers:
                    for col_idx in range(num_cols):
                        cell = tbl.cell(0, col_idx)
                        _set_cell_padding(cell)
                        _shade_cell(cell, "E9EFF7")
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(2)
                        val = headers[col_idx] if col_idx < len(headers) else ""
                        run = p.add_run(str(val))
                        run.bold = True
                    row_cursor = 1

                for r_idx, row_vals in enumerate(rows):
                    for col_idx in range(num_cols):
                        cell = tbl.cell(row_cursor + r_idx, col_idx)
                        _set_cell_padding(cell)
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(2)
                        val = row_vals[col_idx] if col_idx < len(row_vals) else ""
                        p.add_run(str(val))

                spacer = doc.add_paragraph()
                spacer.paragraph_format.space_after = Pt(6)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# FastAPI Routes
# ---------------------------------------------------------------------------


@app.get("/")
async def serve_index() -> FileResponse:
    """Serve the single-file frontend UI."""
    index_path = os.path.join(os.path.dirname(__file__), "static", "index.html")
    if not os.path.exists(index_path):
        raise HTTPException(status_code=500, detail="Frontend static/index.html not found.")
    return FileResponse(index_path)


@app.post("/extract", response_model=MergedDocumentPayload)
async def extract_documents(
    request: Request,
    files: list[UploadFile] = File(...),
    mode: Literal["auto", "excel", "word"] = Form("auto"),
) -> MergedDocumentPayload:
    """Validate, OCR, and merge 1-15 uploaded document images."""
    _check_rate_limit(request)

    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        raise HTTPException(
            status_code=500,
            detail="GEMINI_API_KEY environment variable is not configured.",
        )

    if not files or len(files) > MAX_BATCH_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"Please upload between 1 and {MAX_BATCH_SIZE} images.",
        )

    prepared_images: list[tuple[int, str, bytes]] = []
    failed_pages: list[PageError] = []

    for idx, upload in enumerate(files):
        fname = upload.filename or f"page_{idx + 1}.jpg"
        raw = await upload.read()
        if len(raw) > MAX_FILE_SIZE_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"File '{fname}' exceeds the 10 MB per-image limit.",
            )
        try:
            normalized_bytes = _validate_and_normalize_image(raw)
            prepared_images.append((idx, fname, normalized_bytes))
        except ValueError as exc:
            failed_pages.append(
                PageError(page_index=idx, filename=fname, reason=str(exc))
            )

    if not prepared_images:
        raise HTTPException(
            status_code=422,
            detail={
                "message": "All uploaded images failed validation.",
                "failed_pages": [err.model_dump() for err in failed_pages],
            },
        )

    client = genai.Client(api_key=api_key)
    semaphore = asyncio.Semaphore(GEMINI_CONCURRENCY)

    tasks = [
        _extract_single_page(client, semaphore, idx, fname, img_bytes, mode)
        for idx, fname, img_bytes in prepared_images
    ]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    valid_pages: list[tuple[int, PageExtractionResult]] = []
    for (idx, fname, _), res in zip(prepared_images, results):
        if isinstance(res, Exception):
            logger.error("Page %d (%s) extraction failed: %r", idx + 1, fname, res)
            failed_pages.append(
                PageError(
                    page_index=idx,
                    filename=fname,
                    reason=_format_gemini_error(res),
                )
            )
        else:
            valid_pages.append((idx, res))

    failed_pages.sort(key=lambda e: e.page_index)

    if not valid_pages:
        raise HTTPException(
            status_code=422,
            detail={
                "message": "Vision extraction failed for every page in the batch.",
                "failed_pages": [err.model_dump() for err in failed_pages],
            },
        )

    # Resolve output format
    if mode == "excel":
        output_format: Literal["excel", "word"] = "excel"
    elif mode == "word":
        output_format = "word"
    else:
        tabular_count = sum(1 for _, p in valid_pages if p.layout_type == "tabular")
        prose_count = len(valid_pages) - tabular_count
        output_format = "excel" if tabular_count >= prose_count else "word"

    first_filename = valid_pages[0][1].suggested_filename
    clean_slug = _sanitize_filename(first_filename)

    if output_format == "excel":
        sheets = _merge_excel_pages(valid_pages)
        return MergedDocumentPayload(
            output_format="excel",
            filename=clean_slug,
            pages_processed=len(valid_pages),
            failed_pages=failed_pages,
            sheets=sheets,
            pages=None,
        )

    pages = _merge_word_pages(valid_pages)
    return MergedDocumentPayload(
        output_format="word",
        filename=clean_slug,
        pages_processed=len(valid_pages),
        failed_pages=failed_pages,
        sheets=None,
        pages=pages,
    )


@app.post("/export")
async def export_document(payload: MergedDocumentPayload) -> StreamingResponse:
    """Generate and stream the final .xlsx or .docx file from edited JSON."""
    clean_slug = _sanitize_filename(payload.filename)

    if payload.output_format == "excel":
        buf = build_excel(payload)
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ext = "xlsx"
    else:
        buf = build_word(payload)
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ext = "docx"

    headers = {
        "Content-Disposition": f'attachment; filename="{clean_slug}.{ext}"'
    }
    return StreamingResponse(buf, media_type=media_type, headers=headers)
