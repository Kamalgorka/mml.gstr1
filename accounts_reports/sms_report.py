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


def process_monthly_file(uploaded_file, report_name, od_column_name, output_dir):
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

    df = pd.read_excel(uploaded_file, usecols=all_columns)
    df.columns = [str(c).strip() for c in df.columns]

    status_col = find_column(df.columns, "status")
    od_col = find_column(df.columns, od_column_name)

    # Remove Death cases
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
        "type": "monthly",
        "arrear_free_rows": len(arrear_free_df),
        "arrear_rows": len(arrear_df),
        "total_rows_after_death_removed": len(df),
    }


def read_writeoff_file(uploaded_file, report_name):
    uploaded_file.seek(0)

    header_df = pd.read_excel(uploaded_file, nrows=0)
    all_columns = list(header_df.columns)

    usecols = get_actual_columns(
        all_columns,
        WRITE_OFF_OUTPUT_COLUMNS,
        report_name
    )

    uploaded_file.seek(0)

    df = pd.read_excel(uploaded_file, usecols=usecols)
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {}

    for required_col in WRITE_OFF_OUTPUT_COLUMNS:
        actual_col = find_column(df.columns, required_col)
        rename_map[actual_col] = required_col

    df = df.rename(columns=rename_map)

    return df[WRITE_OFF_OUTPUT_COLUMNS].copy()


def process_writeoff_files(files, output_dir):
    wo_df = read_writeoff_file(
        files["Loan OS Write Off"],
        "Loan OS Write Off"
    )

    wo_il_df = read_writeoff_file(
        files["Loan OS Write Off IL"],
        "Loan OS Write Off IL"
    )

    df = pd.concat([wo_df, wo_il_df], ignore_index=True)

    # Keep only WriteOff cases
    df = df[
        df["status"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("writeoff")
    ].copy()

    if df.empty:
        output_df = pd.DataFrame(columns=WRITE_OFF_OUTPUT_COLUMNS)
    else:
        for col in [
            "od_days",
            "total_arrear",
            "outstanding_principal",
            "outstanding_interest",
        ]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        output_df = (
            df.groupby("cust_id", as_index=False)
            .agg({
                "ZONE": "first",
                "REGION": "first",
                "BRANCH_STATE": "first",
                "status": "first",
                "member_name": "first",
                "mobile_number": "first",
                "od_days": "max",
                "total_arrear": "sum",
                "outstanding_principal": "sum",
                "outstanding_interest": "sum",
            })
        )

        output_df = output_df[
            [
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
        ]

    output_path = os.path.join(output_dir, "Write Off Consolidated.csv")
    output_df.to_csv(output_path, index=False, encoding="utf-8-sig")

    return {
        "report_name": "Write Off Consolidated",
        "type": "writeoff",
        "total_writeoff_rows_after_filter": len(df),
        "unique_cust_id_rows": len(output_df),
    }


def process_sms_report(files, output_dir, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)

    monthly_report_config = {
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
        "Monthly Outstanding SMS Data JLG HUB 5": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 5"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data JLG HUB 6": {
            "file": files["Monthly Outstanding SMS Data JLG HUB 6"],
            "od_col": "max_od_days",
        },
        "Monthly Outstanding SMS Data IL": {
            "file": files["Monthly Outstanding SMS Data IL"],
            "od_col": "max_od_days",
        },
    }

    summary = []

    total_steps = len(monthly_report_config) + 2
    current_step = 0

    for report_name, cfg in monthly_report_config.items():
        current_step += 1

        if progress_callback:
            progress_callback(
                int(((current_step - 1) / total_steps) * 90),
                f"Processing {current_step}/{total_steps}: {report_name}"
            )

        result = process_monthly_file(
            uploaded_file=cfg["file"],
            report_name=report_name,
            od_column_name=cfg["od_col"],
            output_dir=output_dir,
        )

        summary.append(result)

        if progress_callback:
            progress_callback(
                int((current_step / total_steps) * 90),
                f"Completed {current_step}/{total_steps}: {report_name}"
            )

    current_step += 1

    if progress_callback:
        progress_callback(
            int(((current_step - 1) / total_steps) * 90),
            "Processing Write Off Consolidation"
        )

    writeoff_summary = process_writeoff_files(
        files=files,
        output_dir=output_dir
    )

    summary.append(writeoff_summary)

    if progress_callback:
        progress_callback(95, "Creating ZIP file...")

    zip_path = os.path.join(output_dir, "SMS_Report_Output.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(output_dir):
            if file.endswith(".csv"):
                zipf.write(
                    os.path.join(output_dir, file),
                    arcname=file
                )

    if progress_callback:
        progress_callback(100, "Completed.")

    return summary, zip_path
