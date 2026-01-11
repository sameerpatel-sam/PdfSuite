# PdfSuite Web (FastAPI)

Modern web UI for common PDF utilities with fully in-memory processing.

## Features
- Merge PDF (multiple uploads)
- Split PDF by page ranges (ZIP output)
- Compress PDF (basic stream compression)
- PDF to JPG (ZIP of images)
- PDF to Word (text-only, layout not preserved)
- Word to PDF (requires `aspose-words`)

## Requirements
- Python 3.10+
- Poppler installed and on PATH for PDF→JPG (`pdf2image`). On Windows, install from: https://github.com/oschwartz10612/poppler-windows and add `bin` to PATH.
- Optional: `aspose-words` for in-memory DOCX→PDF. Included in requirements; evaluation watermark may apply.

## Setup
```powershell
cd c:\Users\samee\source\repos\PdfSuite\PdfWeb
python -m venv .venv
. .venv\Scripts\Activate.ps1
pip install -r requirements.txt

# run
uvicorn app.main:app --reload
```

Open http://127.0.0.1:8000/

## Notes
- Everything uses `BytesIO`; no files are written to disk.
- PDF compression is conservative (no image re-encoding). For stronger compression, integrate Ghostscript/ImageMagick server-side pipelines.
- PDF→DOCX uses text extraction; formatting/layout will differ from source PDF.
- DOCX→PDF uses Aspose.Words purely in-memory. If unavailable, endpoint returns a clear error.
