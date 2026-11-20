import io, os, re, shutil, subprocess, tempfile, uuid, zipfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from PIL import Image
from pypdf import PdfReader, PdfWriter
import img2pdf
import pytesseract
from pdf2docx import Converter
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024  # 40 MB total request
ALLOWED = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
MAX_FILES = 20

def check_files(files):
    files = [f for f in files if f and f.filename]
    if not files:
        raise ValueError("Please select at least one file.")
    if len(files) > MAX_FILES:
        raise ValueError(f"Please upload no more than {MAX_FILES} files at once.")
    for f in files:
        ext = Path(secure_filename(f.filename)).suffix.lower()
        if ext not in ALLOWED:
            raise ValueError(f"Unsupported file type: {f.filename}")
    return files

def office_convert(src, outdir, target):
    office = shutil.which("libreoffice") or shutil.which("soffice")
    if not office:
        raise RuntimeError("LibreOffice is required for Word/Excel conversion. Install it and ensure 'libreoffice' is on PATH.")
    result = subprocess.run([office, "--headless", "--convert-to", target, "--outdir", str(outdir), str(src)],
                            capture_output=True, text=True, timeout=90)
    expected = outdir / (src.stem + "." + target.split(":")[0])
    if result.returncode != 0 or not expected.exists():
        raise RuntimeError("Office conversion failed. " + (result.stderr or result.stdout or "Check the uploaded file."))
    return expected

def to_pdf(src, outdir):
    ext = src.suffix.lower()
    if ext == ".pdf":
        return src
    if ext in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}:
        out = outdir / (src.stem + ".pdf")
        with open(out, "wb") as fp:
            fp.write(img2pdf.convert(str(src)))
        return out
    if ext in {".doc", ".docx", ".xls", ".xlsx"}:
        return office_convert(src, outdir, "pdf")
    raise ValueError(f"Cannot convert {ext} to PDF.")

def run_conversion(action, uploads, work):
    saved = []
    for f in uploads:
        name = secure_filename(f.filename) or ("upload_" + uuid.uuid4().hex)
        path = work / (uuid.uuid4().hex + "_" + name)
        f.save(path)
        saved.append(path)
    outdir = work / "out"
    outdir.mkdir()
    outputs = []

    if action == "merge":
        writer = PdfWriter()
        for p in saved:
            if p.suffix.lower() != ".pdf": raise ValueError("Merge PDFs accepts PDF files only.")
            reader = PdfReader(str(p))
            for page in reader.pages: writer.add_page(page)
        out = outdir / "merged.pdf"
        with open(out, "wb") as fp: writer.write(fp)
        outputs.append(out)

    elif action == "split":
        for p in saved:
            if p.suffix.lower() != ".pdf": raise ValueError("Split PDFs accepts PDF files only.")
            reader = PdfReader(str(p))
            for i, page in enumerate(reader.pages, 1):
                writer = PdfWriter(); writer.add_page(page)
                out = outdir / f"{p.stem}_page_{i}.pdf"
                with open(out, "wb") as fp: writer.write(fp)
                outputs.append(out)

    elif action == "compress":
        for p in saved:
            if p.suffix.lower() != ".pdf": raise ValueError("Compress PDF accepts PDF files only.")
            reader = PdfReader(str(p)); writer = PdfWriter()
            for page in reader.pages:
                try: page.compress_content_streams()
                except Exception: pass
                writer.add_page(page)
            out = outdir / f"{p.stem}_compressed.pdf"
            with open(out, "wb") as fp: writer.write(fp)
            outputs.append(out)

    elif action == "pdf_word":
        for p in saved:
            if p.suffix.lower() != ".pdf": raise ValueError("PDF to Word accepts PDF files only.")
            out = outdir / f"{p.stem}.docx"
            cv = Converter(str(p))
            try: cv.convert(str(out), start=0, end=None)
            finally: cv.close()
            outputs.append(out)

    elif action == "pdf_excel":
        for p in saved:
            if p.suffix.lower() != ".pdf":
                raise ValueError("PDF to Excel accepts PDF files only.")

            wb = Workbook()
            ws = wb.active
            ws.title = "PDF Data"
            ws.freeze_panes = "A2"
            ws.sheet_view.showGridLines = False

            header_fill = PatternFill("solid", fgColor="1F4E78")
            header_font = Font(bold=True, color="FFFFFF")
            thin_gray = Side(style="thin", color="D9E2F3")
            cell_border = Border(bottom=thin_gray)

            with pdfplumber.open(str(p)) as pdf:
                for page_no, page in enumerate(pdf.pages, 1):
                    # Try common table layouts before falling back to page text.
                    tables = []
                    table_settings_options = [
                        {
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "snap_tolerance": 4,
                            "join_tolerance": 4,
                            "intersection_tolerance": 5,
                        },
                        {
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text",
                            "min_words_vertical": 2,
                            "min_words_horizontal": 1,
                            "text_x_tolerance": 2,
                            "text_y_tolerance": 3,
                        },
                    ]

                    for settings in table_settings_options:
                        try:
                            candidate_tables = page.extract_tables(table_settings=settings) or []
                            candidate_tables = [
                                table for table in candidate_tables
                                if table and any(
                                    any(cell is not None and str(cell).strip() for cell in row)
                                    for row in table
                                )
                            ]
                            if candidate_tables:
                                tables = candidate_tables
                                break
                        except Exception:
                            continue

                    if tables:
                        for table_no, table in enumerate(tables, 1):
                            # Label each extracted table so multiple pages/tables are identifiable.
                            ws.append([f"Page {page_no} — Table {table_no}"])
                            title_row = ws.max_row
                            ws.cell(title_row, 1).font = Font(bold=True, color="1F1F1F", size=11)

                            first_data_row = ws.max_row + 1
                            max_cols = max((len(row) for row in table), default=1)

                            for row in table:
                                cleaned = []
                                for cell in row:
                                    if cell is None:
                                        cleaned.append("")
                                    else:
                                        cleaned.append(re.sub(r"\s+", " ", str(cell)).strip())
                                cleaned.extend([""] * (max_cols - len(cleaned)))
                                ws.append(cleaned)

                            # Style the first row of each detected table as its header.
                            for cell in ws[first_data_row]:
                                cell.fill = header_fill
                                cell.font = header_font
                                cell.alignment = Alignment(vertical="top", wrap_text=True)
                                cell.border = cell_border

                            for row in ws.iter_rows(min_row=first_data_row + 1, max_row=ws.max_row):
                                for cell in row:
                                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                                    cell.border = cell_border
                            ws.append([])
                    else:
                        # Layout mode often preserves column spacing in digitally generated PDFs.
                        text = page.extract_text(layout=True) or page.extract_text() or ""
                        lines = [line.rstrip() for line in text.splitlines() if line.strip()]
                        if not lines:
                            ws.append([f"Page {page_no}: No extractable text found. This may be a scanned PDF; use OCR."])
                            ws.append([])
                            continue

                        ws.append([f"Page {page_no} — extracted text (table structure not detected)"])
                        title_row = ws.max_row
                        ws.cell(title_row, 1).font = Font(bold=True, color="1F1F1F", size=11)

                        first_data_row = ws.max_row + 1
                        for line in lines:
                            # Split only on wider gaps; ordinary spaces inside names/phrases stay together.
                            parts = [part.strip() for part in re.split(r"\s{2,}", line.strip())]
                            ws.append(parts if parts else [line.strip()])

                        for row_idx in range(first_data_row, ws.max_row + 1):
                            for cell in ws[row_idx]:
                                cell.alignment = Alignment(vertical="top", wrap_text=True)
                                cell.border = cell_border
                        ws.append([])

            # Set readable widths without allowing long text to create enormous columns.
            for col_idx in range(1, ws.max_column + 1):
                max_len = 0
                for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=1, max_row=ws.max_row):
                    for item in cell:
                        value = "" if item.value is None else str(item.value)
                        max_len = max(max_len, min(len(value), 55))
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 38)

            out = outdir / f"{p.stem}.xlsx"
            wb.save(out)
            outputs.append(out)

    elif action == "pdf_images":
        for p in saved:
            if p.suffix.lower() != ".pdf": raise ValueError("PDF to JPG/PNG accepts PDF files only.")
            try: import fitz
            except ImportError: raise RuntimeError("PyMuPDF is required for PDF image export.")
            doc = fitz.open(p)
            for i, page in enumerate(doc, 1):
                pix = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
                out = outdir / f"{p.stem}_page_{i}.png"; pix.save(out); outputs.append(out)
            doc.close()

    elif action == "ocr":
        try: import fitz
        except ImportError: raise RuntimeError("PyMuPDF is required for OCR.")
        for p in saved:
            if p.suffix.lower() != ".pdf": raise ValueError("OCR scanned documents currently accepts PDF files.")
            doc = fitz.open(p); chunks = []
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                chunks.append(pytesseract.image_to_string(img))
            doc.close()
            out = outdir / f"{p.stem}_ocr.txt"; out.write_text("\n\n".join(chunks), encoding="utf-8"); outputs.append(out)

    elif action in {"word_pdf", "excel_pdf", "image_pdf"}:
        for p in saved:
            ext = p.suffix.lower()
            if action == "word_pdf" and ext not in {".doc", ".docx"}: raise ValueError("Select Word documents for Word to PDF.")
            if action == "excel_pdf" and ext not in {".xls", ".xlsx"}: raise ValueError("Select Excel documents for Excel to PDF.")
            if action == "image_pdf" and ext not in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}: raise ValueError("Select image files for Image to PDF.")
            outputs.append(to_pdf(p, outdir))

    else:
        raise ValueError("Unknown conversion option.")

    if not outputs: raise RuntimeError("No output was produced.")
    if len(outputs) == 1:
        return outputs[0], outputs[0].name
    bundle = outdir / "docuconvert_results.zip"
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
        for out in outputs: z.write(out, out.name)
    return bundle, bundle.name

@app.get("/")
def home():
    return render_template("index.html")

@app.post("/convert")
def convert():
    action = request.form.get("action", "")
    uploads = request.files.getlist("files")
    try:
        uploads = check_files(uploads)
        # Temporary directory is removed after response is sent, including errors.
        work = Path(tempfile.mkdtemp(prefix="docuconvert_"))
        try:
            output, download_name = run_conversion(action, uploads, work)
            data = output.read_bytes()
            mimetype = "application/zip" if output.suffix == ".zip" else {
                ".pdf":"application/pdf", ".docx":"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xlsx":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".png":"image/png", ".txt":"text/plain"
            }.get(output.suffix, "application/octet-stream")
            response = send_file(io.BytesIO(data), mimetype=mimetype, as_attachment=True, download_name=download_name)
            @response.call_on_close
            def cleanup(): shutil.rmtree(work, ignore_errors=True)
            return response
        except Exception:
            shutil.rmtree(work, ignore_errors=True)
            raise
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "Upload too large. Maximum total upload size is 40 MB."}), 413

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
