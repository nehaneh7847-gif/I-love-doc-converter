# I Love Doc Converter

A Flask-based public document-conversion starter website. Includes a responsive interface, multiple-file uploads, temporary processing, and a 40 MB total request limit.

## Features implemented
- PDF to Word (`pdf2docx`)
- Word to PDF and Excel to PDF (LibreOffice)
- PDF to Excel (extracts detected tables or text lines into a workbook)
- JPG/PNG/TIFF/BMP to PDF
- PDF pages to PNG (one image per page)
- Merge PDFs
- Split PDFs (one output PDF per page)
- Compress PDF (compresses page content streams; may not substantially reduce scanned/image-heavy PDFs)
- OCR scanned PDFs to a downloadable text file (Tesseract)

## Run locally

Requires Python 3.11+ and system dependencies:
- LibreOffice (`libreoffice` on PATH) for Word/Excel to PDF
- Tesseract OCR (`tesseract` on PATH) for OCR

Windows: install LibreOffice and Tesseract OCR, then ensure both are available on PATH.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

## Docker preview

```bash
docker build -t docuconvert .
docker run --rm -p 8080:8080 docuconvert
```
Open http://localhost:8080

## Deploy publicly

1. Push this folder to a private or public GitHub repository.
2. Create a web service on a Docker-capable host such as Render, Railway, or Fly.io.
3. Connect the repository and deploy using the included `Dockerfile` (Render can use `render.yaml`).
4. Confirm the service has enough memory/disk for conversion. PDF/OCR operations can be resource intensive.

Free hosting availability and limits vary; large files or heavy usage may require a paid instance.

## Important production notes
- This is a functional starter, not a security-audited production service.
- Set reverse-proxy request limits and rate limits; add abuse protection, malware scanning, logging with no document contents, and resource/time limits before public launch.
- Temporary workspace is deleted after the response closes, and on handled errors. Unexpected process termination or host failure can leave temporary files behind; production deployments should use a dedicated temp directory with scheduled cleanup.
- Files are processed on the server. Do not claim end-to-end encryption or guaranteed immediate deletion without verifying the hosting environment and transport configuration.
- OCR requires Tesseract. OCR output is plain text, not a searchable PDF.
- PDF-to-Word conversion can alter layout; PDF-to-Excel extraction depends on document structure.
- Compression may have limited effect on scanned PDFs. Add Ghostscript or image recompression for stronger compression.
- For a public service, configure HTTPS, file-type/content validation, per-IP rate limits, anti-malware scanning, and a clear privacy policy before accepting confidential documents.
