# PdfSuite

A simple .NET CLI for common PDF tasks on Windows:
- Merge multiple PDFs
- Split a PDF into single-page files
- Extract text from PDFs

## Web App (FastAPI)
This repo also includes a colorful web app for PDF utilities under [webapp](webapp). It runs on FastAPI and is the recommended UI moving forward.

### Run
```powershell
Push-Location "c:\Users\samee\source\repos\PdfSuite\webapp"
$env:PYTHONPATH=$PWD
& "C:/Users/samee/source/repos/PdfSuite/.venv/Scripts/python.exe" -m uvicorn app.main:app --host 0.0.0.0 --port 8001 --reload
```
Open http://localhost:8001

### Note on PdfWeb
There is a legacy prototype under [PdfWeb](PdfWeb) with a simpler UI and different dependencies (e.g., `pdf2image`). If you are standardizing on the colorful app in [webapp](webapp), you can ignore `PdfWeb/` or remove it to avoid confusion.

## Prerequisites
- .NET SDK 8 or 9 (latest recommended)

## Quick Start
```powershell
# Build
cd c:\Users\samee\source\repos\PdfSuite
 dotnet build

# Show help
 dotnet run --project .\PdfSuite.CLI\PdfSuite.CLI.csproj -- --help
```

## Commands

### merge
Merge multiple PDFs into one.
```powershell
 dotnet run --project .\PdfSuite.CLI\PdfSuite.CLI.csproj -- merge -o out.pdf in1.pdf in2.pdf in3.pdf
```

### split
Split a PDF into single-page PDFs. If `--outdir` is omitted, an `<name>_pages` folder is created next to the input.
```powershell
 dotnet run --project .\PdfSuite.CLI\PdfSuite.CLI.csproj -- split -o .\output-dir input.pdf
```

### extract-text
Extract text from all pages. If `--output` is omitted, text prints to the console.
```powershell
 dotnet run --project .\PdfSuite.CLI\PdfSuite.CLI.csproj -- extract-text -o text.txt input.pdf
```

## Libraries
- PDF merge/split: PDFsharp
- Text extraction: UglyToad.PdfPig (pre-release)
- CLI framework: Spectre.Console / Spectre.Console.Cli

## Notes
- For scanned/image-only PDFs, text extraction requires OCR (not included). Consider integrating Tesseract for OCR in a future command.
- iText7/QuestPDF can be added for advanced generation/templating if needed.