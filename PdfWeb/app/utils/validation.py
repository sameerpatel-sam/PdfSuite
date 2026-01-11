import io
import zipfile

def ensure_pdf(data: bytes) -> None:
    if not data or not data.startswith(b"%PDF"):
        raise ValueError("Invalid PDF file")

def ensure_docx(data: bytes) -> None:
    # DOCX is a zip containing word/ entries
    b = io.BytesIO(data)
    if not zipfile.is_zipfile(b):
        raise ValueError("Invalid DOCX file")
    with zipfile.ZipFile(b) as z:
        names = set(z.namelist())
        if not any(n.startswith("word/") for n in names):
            raise ValueError("Invalid DOCX structure")

# 25 MB limit by default
_MAX_BYTES = 25 * 1024 * 1024

def limit_size(current_bytes: int, max_bytes: int = _MAX_BYTES) -> None:
    if current_bytes > max_bytes:
        raise ValueError("Total upload size exceeds limit")
