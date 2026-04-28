from flask import Flask, request, render_template, send_file, make_response
import pdfplumber
from openpyxl import Workbook, load_workbook
from io import BytesIO
from werkzeug.utils import secure_filename
import os
import re
import tempfile

app = Flask(__name__)


def safe_numeric_eval(formula_text):
    if not isinstance(formula_text, str):
        return formula_text

    text = formula_text.strip()
    if not text.startswith("="):
        return formula_text

    expr = text[1:].strip()

    if not re.fullmatch(r"[0-9\.\+\-\*/\(\)\s]+", expr):
        return formula_text

    try:
        return eval(expr, {"__builtins__": {}}, {})
    except Exception:
        return formula_text


def make_output_filename(original_filename, suffix=" extracted data", new_extension=".xlsx"):
    base_name = os.path.splitext(os.path.basename(original_filename))[0]
    return f"{base_name}{suffix}{new_extension}"


def save_uploaded_file(file_storage, suffix):
    safe_name = secure_filename(file_storage.filename or "upload")
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix="upload_") as tmp:
        file_storage.save(tmp.name)
        return tmp.name, safe_name


def extract_connotes_from_pdf(filepath):
    rows = []

    with pdfplumber.open(filepath) as pdf:
        full_text = ""

        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"

        lines = full_text.splitlines()
        in_table = False

        for line in lines:
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

                # Expected Northline row ending:
                # ... TotalItems Weight Cubic $TotalCost
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
                        cubic_weight = float(raw_cubic) * 250

                        rows.append({
                            "connote": connote,
                            "weight": float(weight),
                            "cubic": cubic_weight,
                            "total_cost": float(total_cost)
                        })

    return rows


def normalize_header(text):
    if text is None:
        return ""
    return str(text).strip().lower().replace(" ", "")


def find_column_indexes(header_row):
    headers = [normalize_header(x) for x in header_row]

    connote_idx = None
    amount_idx = None
    first_comment_idx = None

    for i, header in enumerate(headers):
        if connote_idx is None and header == "connotecode":
            connote_idx = i
        if amount_idx is None and header == "amount":
            amount_idx = i
        if first_comment_idx is None and header == "comment":
            first_comment_idx = i

    return connote_idx, amount_idx, first_comment_idx


def is_valid_connote_code(value):
    if value is None:
        return False

    text = str(value).strip()
    if not text:
        return False

    if not re.fullmatch(r"[A-Z0-9\-]+", text, flags=re.IGNORECASE):
        return False

    if not re.search(r"[A-Z]", text, flags=re.IGNORECASE):
        return False

    if not re.search(r"\d", text):
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

        connote_col = connote_idx + 1
        amount_col = amount_idx + 1
        comment_col = first_comment_idx + 1

        for row_num in range(2, ws_values.max_row + 1):
            connote = ws_values.cell(row_num, connote_col).value
            amount = ws_values.cell(row_num, amount_col).value
            comment = ws_values.cell(row_num, comment_col).value

            if amount is None or str(amount).strip() == "":
                formula_value = ws_formulas.cell(row_num, amount_col).value
                amount = safe_numeric_eval(formula_value)

            if amount is None or str(amount).strip() == "":
                continue

            if not is_valid_connote_code(connote):
                continue

            connote_text = str(connote).strip()
            comment_text = "" if comment is None else str(comment).strip()

            extracted_rows.append([connote_text, amount, comment_text])

    return extracted_rows


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/extract-pdf-connotes", methods=["POST"])
def extract_pdf_connotes():
    files = request.files.getlist("pdf_files")
    remove_duplicates = request.form.get("remove_duplicates") == "on"
    include_cost_report = request.form.get("include_cost_report") == "on"

    try:
        fuel_percent = float(request.form.get("fuel_surcharge_percent") or 0)
    except ValueError:
        fuel_percent = 0

    all_rows = []
    original_names = []

    for file in files:
        if file and file.filename and file.filename.lower().endswith(".pdf"):
            original_names.append(file.filename)
            temp_path, _ = save_uploaded_file(file, ".pdf")

            try:
                rows = extract_connotes_from_pdf(temp_path)
                all_rows.extend(rows)
            finally:
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    if remove_duplicates:
        seen = set()
        unique_rows = []
        for row in all_rows:
            if row["connote"] not in seen:
                seen.add(row["connote"])
                unique_rows.append(row)
        all_rows = unique_rows

    wb = Workbook()
    ws = wb.active
    ws.title = "Connotes"

    if include_cost_report:
        ws.append(["Connote Number", "Weight (Kg)", "Cubic Weight", "Cost Ex Fuel"])
    else:
        ws.append(["Connote Number", "Weight (Kg)", "Cubic Weight"])

    for row in all_rows:
        if include_cost_report:
            if fuel_percent > 0:
                from decimal import Decimal, ROUND_HALF_UP

                cost_ex_fuel = float(
                    (Decimal(str(row["total_cost"])) /
                     (Decimal("1") + Decimal(str(fuel_percent)) / Decimal("100")))
                    .quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                )
            else:
                cost_ex_fuel = row["total_cost"]

            ws.append([row["connote"], row["weight"], row["cubic"], cost_ex_fuel])
        else:
            ws.append([row["connote"], row["weight"], row["cubic"]])

    for cell in ws["B"][1:]:
        cell.number_format = '0.00'

    for cell in ws["C"][1:]:
        cell.number_format = '0.00'

    if include_cost_report:
        for cell in ws["D"][1:]:
            cell.number_format = '$#,##0.00'

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    if len(original_names) == 1:
        download_name = make_output_filename(original_names[0])
    else:
        download_name = "combined extracted data.xlsx"

    response = make_response(send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ))

    response.headers["X-Extracted-Rows"] = str(len(all_rows))
    return response


@app.route("/extract-amount-comments", methods=["POST"])
def extract_amount_comments():
    files = request.files.getlist("excel_files")
    all_rows = []
    original_names = []

    for file in files:
        if file and file.filename:
            filename_lower = file.filename.lower()

            if filename_lower.endswith(".xlsx") or filename_lower.endswith(".xlsm"):
                original_names.append(file.filename)
                suffix = os.path.splitext(file.filename)[1] or ".xlsx"
                temp_path, _ = save_uploaded_file(file, suffix)

                try:
                    rows = extract_amount_rows_from_excel(temp_path)
                    all_rows.extend(rows)
                finally:
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

    wb = Workbook()
    ws = wb.active
    ws.title = "Amount Extract"
    ws.append(["Connote Code", "Amount", "First Comment"])

    for row in all_rows:
        ws.append(row)

    for cell in ws["B"][1:]:
        cell.number_format = '$#,##0.00'

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    if len(original_names) == 1:
        download_name = make_output_filename(original_names[0])
    else:
        download_name = "combined extracted data.xlsx"

    response = make_response(send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ))

    response.headers["X-Extracted-Rows"] = str(len(all_rows))
    return response


if __name__ == "__main__":
    app.run(debug=True)
