import io
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request

from .services.pdf_ops import (
    merge_pdfs,
    split_pdf_by_ranges,
    compress_pdf_basic,
    pdf_to_jpg_zip,
    pdf_to_docx_bytes,
    docx_to_pdf_bytes,
)
from .utils.validation import (
    ensure_pdf,
    ensure_docx,
    limit_size,
)

app = FastAPI(title="PdfSuite Web")

app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/api/merge")
async def api_merge(files: list[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Provide at least two PDF files")
    pdf_buffers = []
    total = 0
    for f in files:
        data = await f.read()
        total += len(data)
        limit_size(total)
        ensure_pdf(data)
        pdf_buffers.append(io.BytesIO(data))
    out = merge_pdfs(pdf_buffers)
    return StreamingResponse(io.BytesIO(out), media_type="application/pdf", headers={
        "Content-Disposition": "attachment; filename=merged.pdf"
    })

@app.post("/api/split")
async def api_split(file: UploadFile = File(...), ranges: str = Form(...)):
    data = await file.read()
    limit_size(len(data))
    ensure_pdf(data)
    out_zip = split_pdf_by_ranges(io.BytesIO(data), ranges)
    return StreamingResponse(io.BytesIO(out_zip), media_type="application/zip", headers={
        "Content-Disposition": "attachment; filename=split.zip"
    })

@app.post("/api/compress")
async def api_compress(file: UploadFile = File(...)):
    data = await file.read()
    limit_size(len(data))
    ensure_pdf(data)
    out = compress_pdf_basic(io.BytesIO(data))
    return StreamingResponse(io.BytesIO(out), media_type="application/pdf", headers={
        "Content-Disposition": "attachment; filename=compressed.pdf"
    })

@app.post("/api/pdf-to-jpg")
async def api_pdf_to_jpg(file: UploadFile = File(...), dpi: int = Form(150), quality: int = Form(85)):
    data = await file.read()
    limit_size(len(data))
    ensure_pdf(data)
    try:
        out_zip = pdf_to_jpg_zip(data, dpi=dpi, quality=quality)
    except RuntimeError as e:
        # Likely missing poppler
        raise HTTPException(status_code=500, detail=str(e))
    return StreamingResponse(io.BytesIO(out_zip), media_type="application/zip", headers={
        "Content-Disposition": "attachment; filename=images.zip"
    })

@app.post("/api/pdf-to-word")
async def api_pdf_to_word(file: UploadFile = File(...)):
    data = await file.read()
    limit_size(len(data))
    ensure_pdf(data)
    out = pdf_to_docx_bytes(data)
    return StreamingResponse(io.BytesIO(out), media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", headers={
        "Content-Disposition": "attachment; filename=converted.docx"
    })

@app.post("/api/word-to-pdf")
async def api_word_to_pdf(file: UploadFile = File(...)):
    data = await file.read()
    limit_size(len(data))
    ensure_docx(data)
    try:
        out = docx_to_pdf_bytes(data)
    except ImportError:
        raise HTTPException(status_code=500, detail="Word to PDF requires 'aspose-words' installed. Please install the dependency or use a host with it available.")
    return StreamingResponse(io.BytesIO(out), media_type="application/pdf", headers={
        "Content-Disposition": "attachment; filename=converted.pdf"
    })
