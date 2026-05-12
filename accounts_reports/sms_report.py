import os
import re
import zipfile
from openpyxl import load_workbook, Workbook


WRITE_OFF_OUTPUT_COLUMNS = [
    "ZONE",
    "REGION",
    "BRANCH_STATE",
    "status",
    "cust_id",
    "member_name",
    "mobile_number",
    "od_days",
    "total_arrear",
    "outstanding_principal",
    "outstanding_interest",
]


def clean_file_name(name):
    name = str(name).replace(".xlsx", "").replace(".xls", "")
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    return name.strip()


def norm_col(col):
    return str(col).strip().lower().replace(" ", "").replace("_", "")


def to_number(value):
    try:
        if value is None or str(value).strip() == "":
            return 0
        return float(value)
    except Exception:
        return 0


def save_uploaded_file(uploaded_file, output_dir, report_name):
    raw_path = os.path.join(output_dir, clean_file_name(report_name) + "_RAW.xlsx")
    uploaded_file.seek(0)

    with open(raw_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return raw_path


def split_arrear_files(
    uploaded_file,
    report_name,
    od_column_name,
    output_dir,
    writeoff_only=False
):
    raw_path = save_uploaded_file(uploaded_file, output_dir, report_name)

    wb_in = load_workbook(raw_path, read_only=True, data_only=True)
    ws_in = wb_in.active

    header_row = next(ws_in.iter_rows(min_row=1, max_row=1, values_only=True))
    headers = list(header_row)
    norm_headers = [norm_col(h) for h in headers]

    status_key = norm_col("status")
    od_key = norm_col(od_column_name)

    if status_key not in norm_headers:
        raise ValueError(f"Status column not found in {report_name}")

    if od_key not in norm_headers:
        raise ValueError(f"{od_column_name} column not found in {report_name}")

    status_idx = norm_headers.index(status_key)
    od_idx = norm_headers.index(od_key)

    if writeoff_only:
        output_indexes = []
        output_headers = []

        for col_name in WRITE_OFF_OUTPUT_COLUMNS:
            col_key = norm_col(col_name)

            if col_key not in norm_headers:
                raise ValueError(f"Column '{col_name}' not found in {report_name}")

            idx = norm_headers.index(col_key)
            output_indexes.append(idx)
            output_headers.append(headers[idx])
    else:
        output_indexes = list(range(len(headers)))
        output_headers = headers

    safe_report_name = clean_file_name(report_name)

    arrear_free_name = f"Arrear Free {safe_report_name}.xlsx"
    arrear_name = f"Arrear {safe_report_name}.xlsx"

    arrear_free_path = os.path.join(output_dir, arrear_free_name)
    arrear_path = os.path.join(output_dir, arrear_name)

    wb_free = Workbook(write_only=True)
    ws_free = wb_free.create_sheet("Data")
    ws_free.append(output_headers)

    wb_arrear = Workbook(write_only=True)
    ws_arrear = wb_arrear.create_sheet("Data")
    ws_arrear.append(output_headers)

    total_rows = 0
    arrear_free_rows = 0
    arrear_rows = 0

    for row in ws_in.iter_rows(min_row=2, values_only=True):
        status_value = str(row[status_idx] or "").strip().lower()

        if "death" in status_value:
            continue

        od_value = to_number(row[od_idx])
        output_row = [row[i] if i < len(row) else None for i in output_indexes]

        total_rows += 1

        if od_value == 0:
            ws_free.append(output_row)
            arrear_free_rows += 1

        elif od_value > 0:
            ws_arrear.append(output_row)
            arrear_rows += 1

    wb_free.save(arrear_free_path)
    wb_arrear.save(arrear_path)

    wb_in.close()

    try:
        os.remove(raw_path)
    except Exception:
        pass

    return {
        "report_name": safe_report_name,
        "arrear_free_file": arrear_free_name,
        "arrear_file": arrear_name,
        "total_rows_after_death_removed": total_rows,
        "arrear_free_rows": arrear_free_rows,
        "arrear_rows": arrear_rows,
    }


def process_sms_report(files, output_dir, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)

    report_config = {
        "Monthly Outstanding SMS Data JLG HUB 1": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 1"],
            "od_col": "max_od_days",
            "writeoff_only": False,
        },
        "Monthly Outstanding SMS Data JLG HUB 2": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 2"],
            "od_col": "max_od_days",
            "writeoff_only": False,
        },
        "Monthly Outstanding SMS Data JLG HUB 3": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 3"],
            "od_col": "max_od_days",
            "writeoff_only": False,
        },
        "Monthly Outstanding SMS Data JLG HUB 4": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 4"],
            "od_col": "max_od_days",
            "writeoff_only": False,
        },
        "Monthly Outstanding SMS Data JLG HUB 5 and 6": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 5 and 6"],
            "od_col": "max_od_days",
            "writeoff_only": False,
        },
        "Monthly Outstanding SMS Data IL": {
            "file": files["Monthly Outstanding SMS Data IL"],
            "od_col": "max_od_days",
            "writeoff_only": False,
        },
        "Loan OS Write Off": {
            "file": files["Loan OS Write Off"],
            "od_col": "od_days",
            "writeoff_only": True,
        },
        "Loan OS Write Off IL": {
            "file": files["Loan OS Write Off IL"],
            "od_col": "od_days",
            "writeoff_only": True,
        },
    }

    summary = []
    total_reports = len(report_config)

    for idx, (report_name, cfg) in enumerate(report_config.items(), start=1):

        if progress_callback:
            progress_callback(
                int(((idx - 1) / total_reports) * 90),
                f"Processing {idx}/{total_reports}: {report_name}"
            )

        result = split_arrear_files(
            uploaded_file=cfg["file"],
            report_name=report_name,
            od_column_name=cfg["od_col"],
            output_dir=output_dir,
            writeoff_only=cfg["writeoff_only"],
        )

        summary.append(result)

        if progress_callback:
            progress_callback(
                int((idx / total_reports) * 90),
                f"Completed {idx}/{total_reports}: {report_name}"
            )

    if progress_callback:
        progress_callback(95, "Creating ZIP file...")

    zip_path = os.path.join(output_dir, "SMS_Report_Output.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(output_dir):
            if file.endswith(".xlsx"):
                zipf.write(
                    os.path.join(output_dir, file),
                    arcname=file
                )

    if progress_callback:
        progress_callback(100, "Completed.")

    return summary, zip_path
