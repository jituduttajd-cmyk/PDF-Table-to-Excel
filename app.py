import io
import re
from pathlib import Path

import pdfplumber
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024
CORS(app)

HEADING_FILL = PatternFill(start_color="FF2563EB", end_color="FF2563EB", fill_type="solid")
HEADING_FONT = Font(bold=True, color="FFFFFFFF")
THIN = Side(style="thin", color="FFD1D5DB")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
WRAP_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")
NUMERIC_RE = re.compile(r"^-?\d[\d,]*\.?\d*$")

def clean_cell(text):
    if text is None:
        return ""
    return str(text).strip()

def is_english_row(row):
    joined = " ".join(c for c in row if c)
    visible = re.sub(r"\s", "", joined)
    if not visible:
        return False
    ascii_chars = sum(1 for ch in visible if ord(ch) < 128)
    return ascii_chars / len(visible) >= 0.60

def break_after_comma(text):
    if text is None:
        return ""
    s = str(text).strip()
    if not s or NUMERIC_RE.match(s):
        return s
    return ",\n".join(p.strip() for p in s.split(",") if p.strip())

def normalize_row(row, ncols):
    row = [clean_cell(c) for c in row]
    if len(row) < ncols:
        row += [""] * (ncols - len(row))
    return row[:ncols]

def looks_like_header_block(rows):
    if len(rows) < 2:
        return False
    second = rows[1]
    cells_ok = all(re.match(r"^\(\d+\)$", c.strip()) or c.strip() == "" for c in second)
    return cells_ok and any(c.strip() for c in second)

def merge_continuation(current_table, row):
    first_blank = row[0].strip() == ""
    has_other = any(c.strip() for c in row[1:])
    if first_blank and has_other and len(current_table) > 1:
        prev = current_table[-1]
        merged = []
        for a, b in zip(prev, row):
            if b.strip():
                merged.append((a + "\n" + b).strip() if a.strip() else b)
            else:
                merged.append(a)
        current_table[-1] = merged
        return True
    return False

def extract_english_tables(pdf_stream, english_only=True, merge_continuations=True):
    logical_tables = []
    current_table = None
    current_is_english = False

    with pdfplumber.open(pdf_stream) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            for raw_table in page.extract_tables():
                if not raw_table:
                    continue

                ncols = max(len(r) for r in raw_table)
                rows = [normalize_row(r, ncols) for r in raw_table]

                if looks_like_header_block(rows):
                    if current_table:
                        logical_tables.append(current_table)
                    header_row = rows[0]
                    current_is_english = is_english_row(header_row)
                    current_table = [header_row] if (current_is_english or not english_only) else None
                    data_rows = rows[2:]
                else:
                    if current_table is None or (english_only and not current_is_english):
                        continue
                    data_rows = rows

                if current_table is None:
                    continue

                for row in data_rows:
                    if merge_continuations and merge_continuation(current_table, row):
                        continue
                    if any(c.strip() for c in row):
                        current_table.append(row)

    if current_table:
        logical_tables.append(current_table)

    return logical_tables

def write_workbook(tables, comma_breaks=True):
    wb = Workbook()
    wb.remove(wb.active)

    for idx, table in enumerate(tables, start=1):
        sheet_name = f"Table {idx}" if len(tables) > 1 else "Table"
        ws = wb.create_sheet(sheet_name[:31])
        header = table[0]
        ncols = len(header)

        ws.append(header)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = HEADING_FILL
            cell.font = HEADING_FONT
            cell.alignment = WRAP_CENTER
            cell.border = CELL_BORDER

        for row in table[1:]:
            values = [break_after_comma(c) if comma_breaks else clean_cell(c) for c in row]
            ws.append(values)
            r = ws.max_row
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                cell.alignment = WRAP
                cell.border = CELL_BORDER

        for c in range(1, ncols + 1):
            col_letter = get_column_letter(c)
            max_len = 10
            for r in range(1, ws.max_row + 1):
                val = ws.cell(r, c).value or ""
                for line in str(val).split("\n"):
                    max_len = max(max_len, len(line))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 30

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.get("/health")
def health():
    return jsonify({"status":"ok","engine":"pdfplumber+openpyxl"})

@app.post("/convert")
def convert():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename.lower().endswith(".pdf"):
        return jsonify({"error":"Please upload a PDF file."}), 400

    try:
        data = uploaded.read()
        if not data:
            return jsonify({"error":"The uploaded PDF is empty."}), 400

        english_only = request.form.get("english_only","1") == "1"
        merge_continuations = request.form.get("merge_continuations","1") == "1"
        comma_breaks = request.form.get("comma_breaks","1") == "1"

        tables = extract_english_tables(
            io.BytesIO(data),
            english_only=english_only,
            merge_continuations=merge_continuations
        )

        if not tables:
            return jsonify({
                "error":"No suitable table was found. The PDF may be scanned/image-only or may not contain ruled tables."
            }), 422

        output = write_workbook(tables, comma_breaks=comma_breaks)
        base = Path(uploaded.filename).stem
        filename = f"{base}_converted.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as exc:
        app.logger.exception("Conversion error")
        return jsonify({"error":f"Conversion failed: {exc}"}), 500

@app.errorhandler(413)
def too_large(_):
    return jsonify({"error":"PDF is too large. Maximum size is 30 MB."}), 413

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
