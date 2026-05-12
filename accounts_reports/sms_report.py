import os
import re
import zipfile
import pandas as pd


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


def find_column(columns, required_col):
    target = norm_col(required_col)
    for col in columns:
        if norm_col(col) == target:
            return col
    return None


def get_actual_columns(columns, required_cols, report_name):
    actual_cols = []
    for col in required_cols:
        actual = find_column(columns, col)
        if actual is None:
            raise ValueError(f"Column '{col}' not found in {report_name}")
        actual_cols.append(actual)
    return actual_cols


def process_single_file(uploaded_file, report_name, od_column_name, output_dir, writeoff_only=False):
    uploaded_file.seek(0)

    header_df = pd.read_excel(uploaded_file, nrows=0)
    all_columns = list(header_df.columns)

    status_col = find_column(all_columns, "status")
    od_col = find_column(all_columns, od_column_name)

    if status_col is None:
        raise ValueError(f"Status column not found in {report_name}")

    if od_col is None:
        raise ValueError(f"{od_column_name} column not found in {report_name}")

    uploaded_file.seek(0)

    if writeoff_only:
        usecols = get_actual_columns(all_columns, WRITE_OFF_OUTPUT_COLUMNS, report_name)
    else:
        usecols = all_columns

    df = pd.read_excel(uploaded_file, usecols=usecols)
    df.columns = [str(c).strip() for c in df.columns]

    status_col = find_column(df.columns, "status")
    od_col = find_column(df.columns, od_column_name)

    df = df[
        ~df[status_col]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.contains("death", na=False)
    ].copy()

    df[od_col] = pd.to_numeric(df[od_col], errors="coerce").fillna(0)

    arrear_free_df = df[df[od_col] == 0].copy()
    arrear_df = df[df[od_col] > 0].copy()

    safe_name = clean_file_name(report_name)

    arrear_free_name = f"Arrear Free {safe_name}.csv"
    arrear_name = f"Arrear {safe_name}.csv"

    arrear_free_path = os.path.join(output_dir, arrear_free_name)
    arrear_path = os.path.join(output_dir, arrear_name)

    arrear_free_df.to_csv(arrear_free_path, index=False, encoding="utf-8-sig")
    arrear_df.to_csv(arrear_path, index=False, encoding="utf-8-sig")

    return {
        "report_name": safe_name,
        "arrear_free_rows": len(arrear_free_df),
        "arrear_rows": len(arrear_df),
        "total_rows_after_death_removed": len(df),
    }


def process_sms_report(files, output_dir, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)

    report_config = {
        "Monthly Outstanding SMS Data JLG HUB 1": {"file": files["Monthly Outstanding SMS Data JLG HUB 1"], "od_col": "max_od_days", "writeoff_only": False},
        "Monthly Outstanding SMS Data JLG HUB 2": {"file": files["Monthly Outstanding SMS Data JLG HUB 2"], "od_col": "max_od_days", "writeoff_only": False},
        "Monthly Outstanding SMS Data JLG HUB 3": {"file": files["Monthly Outstanding SMS Data JLG HUB 3"], "od_col": "max_od_days", "writeoff_only": False},
        "Monthly Outstanding SMS Data JLG HUB 4": {"file": files["Monthly Outstanding SMS Data JLG HUB 4"], "od_col": "max_od_days", "writeoff_only": False},
        "Monthly Outstanding SMS Data JLG HUB 5": {"file": files["Monthly Outstanding SMS Data JLG HUB 5"], "od_col": "max_od_days", "writeoff_only": False},
        "Monthly Outstanding SMS Data JLG HUB 6": {"file": files["Monthly Outstanding SMS Data JLG HUB 6"], "od_col": "max_od_days", "writeoff_only": False},
        "Monthly Outstanding SMS Data IL": {"file": files["Monthly Outstanding SMS Data IL"], "od_col": "max_od_days", "writeoff_only": False},
        "Loan OS Write Off": {"file": files["Loan OS Write Off"], "od_col": "od_days", "writeoff_only": True},
        "Loan OS Write Off IL": {"file": files["Loan OS Write Off IL"], "od_col": "od_days", "writeoff_only": True},
    }

    summary = []
    total = len(report_config)

    for idx, (report_name, cfg) in enumerate(report_config.items(), start=1):
        if progress_callback:
            progress_callback(
                int(((idx - 1) / total) * 90),
                f"Processing {idx}/{total}: {report_name}"
            )

        result = process_single_file(
            uploaded_file=cfg["file"],
            report_name=report_name,
            od_column_name=cfg["od_col"],
            output_dir=output_dir,
            writeoff_only=cfg["writeoff_only"],
        )

        summary.append(result)

        if progress_callback:
            progress_callback(
                int((idx / total) * 90),
                f"Completed {idx}/{total}: {report_name}"
            )

    if progress_callback:
        progress_callback(95, "Creating ZIP file...")

    zip_path = os.path.join(output_dir, "SMS_Report_Output.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(output_dir):
            if file.endswith(".csv"):
                zipf.write(os.path.join(output_dir, file), arcname=file)

    if progress_callback:
        progress_callback(100, "Completed.")

    return summary, zip_path
