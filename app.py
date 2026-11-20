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
            ws.title = "Extracted Data"
            ws.sheet_view.showGridLines = False
            ws.freeze_panes = "A2"
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_view.zoomScale = 90

            header_fill = PatternFill("solid", fgColor="183B66")
            header_font = Font(bold=True, color="FFFFFF", size=10)
            section_fill = PatternFill("solid", fgColor="EAF1F8")
            section_font = Font(bold=True, color="183B66", size=10)
            body_font = Font(name="Aptos", size=10, color="243247")
            bottom_line = Side(style="thin", color="D9E2EF")
            table_border = Border(bottom=bottom_line)

            def clean_cell(value):
                if value is None:
                    return ""
                value = str(value).replace("\x00", " ")
                return re.sub(r"\s+", " ", value).strip()

            def nonempty_rows(table):
                rows = []
                max_cols = max((len(row) for row in table), default=0)
                for row in table:
                    cleaned = [clean_cell(cell) for cell in row]
                    if not any(cleaned):
                        continue
                    cleaned.extend([""] * (max_cols - len(cleaned)))
                    # Skip repeated empty/near-empty extraction artefacts.
                    if rows and cleaned == rows[-1]:
                        continue
                    rows.append(cleaned)
                return rows

            def append_table(table, section_label=None):
                rows = nonempty_rows(table)
                if not rows:
                    return
                if section_label:
                    if ws.max_row > 1 or ws.cell(1, 1).value is not None:
                        ws.append([])
                    ws.append([section_label])
                    section_row = ws.max_row
                    ws.cell(section_row, 1).font = section_font
                    ws.cell(section_row, 1).fill = section_fill
                    ws.merge_cells(start_row=section_row, start_column=1,
                                   end_row=section_row, end_column=max(1, max(len(r) for r in rows)))
                header_row = ws.max_row + 1
                for row in rows:
                    ws.append(row)
                for cell in ws[header_row]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.border = table_border
                ws.row_dimensions[header_row].height = 30
                for row_idx in range(header_row + 1, ws.max_row + 1):
                    for cell in ws[row_idx]:
                        cell.font = body_font
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                        cell.border = table_border
                        if isinstance(cell.value, str):
                            cell.value = cell.value.strip()

            with pdfplumber.open(str(p)) as pdf:
                for page_no, page in enumerate(pdf.pages, 1):
                    tables = []
                    # Try ruled-table extraction first, then spacing/text-based detection.
                    settings_options = [
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
                    for settings in settings_options:
                        try:
                            candidates = page.extract_tables(table_settings=settings) or []
                            candidates = [nonempty_rows(t) for t in candidates if t]
                            candidates = [t for t in candidates if t and max(map(len, t), default=0) >= 2]
                            if candidates:
                                # Prefer the strategy that returns the most useful table cells.
                                tables = candidates
                                break
                        except Exception:
                            continue

                    if tables:
                        for table_no, table in enumerate(tables, 1):
                            append_table(table, f"Page {page_no} · Table {table_no}")
                    else:
                        # Fallback: preserve readable text and split columns only at wide gaps.
                        text = page.extract_text(layout=True) or page.extract_text() or ""
                        lines = [line.rstrip() for line in text.splitlines() if line.strip()]
                        if not lines:
                            append_table([[f"No selectable text found on page {page_no}. This may be a scanned PDF; try OCR first."]],
                                         f"Page {page_no} · No table detected")
                            continue
                        fallback_rows = []
                        for line in lines:
                            parts = [part.strip() for part in re.split(r"\s{3,}", line.strip())]
                            fallback_rows.append(parts if parts else [line.strip()])
                        append_table(fallback_rows, f"Page {page_no} · Text fallback (table layout not detected)")

            # Make the workbook readable without creating extremely wide columns.
            for col_idx in range(1, ws.max_column + 1):
                values = []
                for row_idx in range(1, ws.max_row + 1):
                    value = ws.cell(row_idx, col_idx).value
                    if value is not None:
                        values.append(str(value))
                max_len = max((min(len(v), 48) for v in values), default=10)
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 3, 12), 34)

            ws.auto_filter.ref = f"A1:{get_column_letter(max(1, ws.max_column))}{max(1, ws.max_row)}"
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
        return jsonify({"error": str(e) or "Conversion failed. Please check the file and try again."}), 400

@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "Upload too large. Maximum total upload size is 40 MB."}), 413

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
