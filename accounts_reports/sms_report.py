import os
import re
import zipfile
import pandas as pd


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


def get_required_columns(uploaded_file, od_column_name, report_name):
    header_df = pd.read_excel(uploaded_file, nrows=0)
    columns = list(header_df.columns)

    status_col = find_column(columns, "status")
    od_col = find_column(columns, od_column_name)

    if status_col is None:
        raise ValueError(f"Status column not found in {report_name}")

    if od_col is None:
        raise ValueError(f"{od_column_name} column not found in {report_name}")

    uploaded_file.seek(0)
    return status_col, od_col


def split_arrear_files(uploaded_file, report_name, od_column_name, output_dir):
    status_col, od_col = get_required_columns(
        uploaded_file,
        od_column_name,
        report_name
    )

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        usecols=[status_col, od_col]
    )

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

    safe_report_name = clean_file_name(report_name)

    arrear_free_name = f"Arrear Free {safe_report_name}.xlsx"
    arrear_name = f"Arrear {safe_report_name}.xlsx"

    arrear_free_path = os.path.join(output_dir, arrear_free_name)
    arrear_path = os.path.join(output_dir, arrear_name)

    arrear_free_df.to_excel(arrear_free_path, index=False)
    arrear_df.to_excel(arrear_path, index=False)

    return {
        "report_name": safe_report_name,
        "arrear_free_file": arrear_free_name,
        "arrear_file": arrear_name,
        "total_rows_after_death_removed": len(df),
        "arrear_free_rows": len(arrear_free_df),
        "arrear_rows": len(arrear_df),
    }


def process_sms_report(files, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    report_config = {
        "Monthly Outstanding SMS Data JLG HUB 1": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 1"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data JLG HUB 2": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 2"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data JLG HUB 3": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 3"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data JLG HUB 4": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 4"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data JLG HUB 5 and 6": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 5 and 6"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data IL": {
            "file": files["Monthly Outstanding SMS Data IL"],
            "od_col": "max_od_days",
        },
        "Loan OS Write Off": {
            "file": files["Loan OS Write Off"],
            "od_col": "od_days",
        },
        "Loan OS Write Off IL": {
            "file": files["Loan OS Write Off IL"],
            "od_col": "od_days",
        },
    }

    summary = []

    for report_name, cfg in report_config.items():
        result = split_arrear_files(
            uploaded_file=cfg["file"],
            report_name=report_name,
            od_column_name=cfg["od_col"],
            output_dir=output_dir,
        )
        summary.append(result)

    zip_path = os.path.join(output_dir, "SMS_Report_Output.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(output_dir):
            if file.endswith(".xlsx"):
                zipf.write(
                    os.path.join(output_dir, file),
                    arcname=file
                )

    return summary, zip_path
