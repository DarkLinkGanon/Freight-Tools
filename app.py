
from flask import Flask, request, render_template, send_file, make_response, jsonify
import pdfplumber, os, re, tempfile, base64
from openpyxl import Workbook, load_workbook
from io import BytesIO
from werkzeug.utils import secure_filename
from decimal import Decimal, ROUND_HALF_UP

app = Flask(__name__)

def make_output_filename(original_filename, suffix=" extracted data", new_extension=".xlsx"):
    base_name = os.path.splitext(os.path.basename(original_filename))[0]
    return f"{base_name}{suffix}{new_extension}"

def save_uploaded_file(file_storage, suffix):
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix="upload_") as tmp:
        file_storage.save(tmp.name)
        return tmp.name

def save_base64_pdf(base64_text):
    if "," in base64_text:
        base64_text = base64_text.split(",", 1)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", prefix="upload_json_") as tmp:
        tmp.write(base64.b64decode(base64_text))
        return tmp.name

def money_ex_fuel(total_cost, fuel_percent):
    if fuel_percent <= 0:
        return float(Decimal(str(total_cost)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    return float((Decimal(str(total_cost)) / (Decimal("1") + Decimal(str(fuel_percent)) / Decimal("100"))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def extract_connotes_from_pdf(filepath):
    rows = []
    with pdfplumber.open(filepath) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

    in_table = False
    for line in full_text.splitlines():
        if "Customer Con Note Sender Name Receiver Name" in line:
            in_table = True
            continue
        if in_table:
            if "Manifest Notes:" in line or "Total Connotes:" in line:
                break
            line = line.strip()
            if not line:
                continue
            parts = line.split()
            if len(parts) >= 5:
                connote = parts[1].strip()
                weight = parts[-3].replace(",", "").strip()
                raw_cubic = parts[-2].replace(",", "").strip()
                total_cost = parts[-1].replace("$", "").replace(",", "").strip()
                if (
                    re.fullmatch(r"[A-Z0-9\-]{6,30}", connote)
                    and re.fullmatch(r"\d+(\.\d+)?", weight)
                    and re.fullmatch(r"\d+(\.\d+)?", raw_cubic)
                    and re.fullmatch(r"\d+(\.\d+)?", total_cost)
                ):
                    rows.append({
                        "connote": connote,
                        "weight": float(weight),
                        "cubic": float(raw_cubic) * 250,
                        "total_cost": float(total_cost)
                    })
    return rows

def build_northline_workbook(all_rows, include_cost_report, fuel_percent):
    wb = Workbook()
    ws = wb.active
    ws.title = "Connotes"
    if include_cost_report:
        ws.append(["Connote Number", "Weight (Kg)", "Cubic Weight", "Cost Ex Fuel"])
    else:
        ws.append(["Connote Number", "Weight (Kg)", "Cubic Weight"])
    for row in all_rows:
        if include_cost_report:
            ws.append([row["connote"], row["weight"], row["cubic"], money_ex_fuel(row["total_cost"], fuel_percent)])
        else:
            ws.append([row["connote"], row["weight"], row["cubic"]])
    for cell in ws["B"][1:]:
        cell.number_format = "0.00"
    for cell in ws["C"][1:]:
        cell.number_format = "0.00"
    if include_cost_report:
        for cell in ws["D"][1:]:
            cell.number_format = "$#,##0.00"
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def safe_numeric_eval(formula_text):
    if not isinstance(formula_text, str) or not formula_text.strip().startswith("="):
        return formula_text
    expr = formula_text.strip()[1:].strip()
    if not re.fullmatch(r"[0-9\.\+\-\*/\(\)\s]+", expr):
        return formula_text
    try:
        return eval(expr, {"__builtins__": {}}, {})
    except Exception:
        return formula_text

def normalize_header(text):
    return "" if text is None else str(text).strip().lower().replace(" ", "")

def find_column_indexes(header_row):
    headers = [normalize_header(x) for x in header_row]
    connote_idx = amount_idx = first_comment_idx = None
    for i, header in enumerate(headers):
        if connote_idx is None and header == "connotecode": connote_idx = i
        if amount_idx is None and header == "amount": amount_idx = i
        if first_comment_idx is None and header == "comment": first_comment_idx = i
    return connote_idx, amount_idx, first_comment_idx

def is_valid_connote_code(value):
    if value is None:
        return False

    text = str(value).strip()

    if not text:
        return False

    # Remove .0 if Excel stores whole-number connotes as numbers
    if text.endswith(".0"):
        text = text[:-2]

    # Allow letters, numbers, and hyphens
    if not re.fullmatch(r"[A-Z0-9\-]+", text, flags=re.IGNORECASE):
        return False

    # Must contain at least one number
    if not re.search(r"\d", text):
        return False

    # Must be long enough to be a real connote
    if len(text) < 6:
        return False

    return True

def extract_amount_rows_from_excel(filepath):
    extracted_rows = []
    wb_values = load_workbook(filepath, data_only=True)
    wb_formulas = load_workbook(filepath, data_only=False)
    for sheet_name in wb_values.sheetnames:
        ws_values = wb_values[sheet_name]
        ws_formulas = wb_formulas[sheet_name]
        if ws_values.max_row < 2:
            continue
        header_row = [ws_values.cell(1, c).value for c in range(1, ws_values.max_column + 1)]
        connote_idx, amount_idx, first_comment_idx = find_column_indexes(header_row)
        if connote_idx is None or amount_idx is None or first_comment_idx is None:
            continue
        for row_num in range(2, ws_values.max_row + 1):
            connote = ws_values.cell(row_num, connote_idx + 1).value
            amount = ws_values.cell(row_num, amount_idx + 1).value
            comment = ws_values.cell(row_num, first_comment_idx + 1).value
            if amount is None or str(amount).strip() == "":
                amount = safe_numeric_eval(ws_formulas.cell(row_num, amount_idx + 1).value)
            if amount is None or str(amount).strip() == "" or not is_valid_connote_code(connote):
                continue
            extracted_rows.append([str(connote).strip(), amount, "" if comment is None else str(comment).strip()])
    return extracted_rows

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/extract-pdf-connotes-json", methods=["POST"])
def extract_pdf_connotes_json():
    try:
        payload = request.get_json(force=True)
        files = payload.get("files", [])
        remove_duplicates = bool(payload.get("remove_duplicates", False))
        include_cost_report = bool(payload.get("include_cost_report", False))
        try:
            fuel_percent = float(payload.get("fuel_surcharge_percent") or 0)
        except ValueError:
            fuel_percent = 0

        all_rows, original_names = [], []
        for item in files:
            filename = item.get("name", "upload.pdf")
            content = item.get("content", "")
            if not filename.lower().endswith(".pdf"):
                continue
            original_names.append(filename)
            temp_path = save_base64_pdf(content)
            try:
                all_rows.extend(extract_connotes_from_pdf(temp_path))
            finally:
                try: os.remove(temp_path)
                except Exception: pass

        if remove_duplicates:
            seen, unique_rows = set(), []
            for row in all_rows:
                if row["connote"] not in seen:
                    seen.add(row["connote"])
                    unique_rows.append(row)
            all_rows = unique_rows

        output = build_northline_workbook(all_rows, include_cost_report, fuel_percent)
        download_name = make_output_filename(original_names[0]) if len(original_names) == 1 else "combined extracted data.xlsx"
        response = make_response(send_file(output, as_attachment=True, download_name=download_name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
        response.headers["X-Extracted-Rows"] = str(len(all_rows))
        return response
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/extract-amount-comments", methods=["POST"])
def extract_amount_comments():
    files = request.files.getlist("excel_files")
    all_rows, original_names = [], []
    for file in files:
        if file and file.filename and file.filename.lower().endswith((".xlsx", ".xlsm")):
            original_names.append(file.filename)
            suffix = os.path.splitext(file.filename)[1] or ".xlsx"
            temp_path = save_uploaded_file(file, suffix)
            try:
                all_rows.extend(extract_amount_rows_from_excel(temp_path))
            finally:
                try: os.remove(temp_path)
                except Exception: pass
    wb = Workbook()
    ws = wb.active
    ws.title = "Amount Extract"
    ws.append(["Connote Code", "Amount", "First Comment"])
    for row in all_rows:
        ws.append(row)
    for cell in ws["B"][1:]:
        cell.number_format = "$#,##0.00"
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    download_name = make_output_filename(original_names[0]) if len(original_names) == 1 else "combined extracted data.xlsx"
    response = make_response(send_file(output, as_attachment=True, download_name=download_name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    response.headers["X-Extracted-Rows"] = str(len(all_rows))
    return response

if __name__ == "__main__":
    app.run(debug=True)
