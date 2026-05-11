import os
import re
import zipfile
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def clean_file_name(name):
    name = str(name).replace(".xlsx", "").replace(".xls", "")
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    return name.strip()


def find_column(df, required_col):
    target = required_col.strip().lower().replace(" ", "").replace("_", "")
    for col in df.columns:
        current = str(col).strip().lower().replace(" ", "").replace("_", "")
        if current == target:
            return col
    return None


def format_excel_file(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    header_fill = PatternFill("solid", fgColor="BDD7EE")
    header_font = Font(bold=True, color="000000")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

    wb.save(file_path)


def split_arrear_files(uploaded_file, report_name, od_column_name, output_dir):
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]

    status_col = find_column(df, "status")
    od_col = find_column(df, od_column_name)

    if status_col is None:
        raise ValueError(f"Status column not found in {report_name}")

    if od_col is None:
        raise ValueError(f"{od_column_name} column not found in {report_name}")

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

    safe_report_name = clean_file_name(report_name)

    arrear_free_name = f"Arrear Free {safe_report_name}.xlsx"
    arrear_name = f"Arrear {safe_report_name}.xlsx"

    arrear_free_path = os.path.join(output_dir, arrear_free_name)
    arrear_path = os.path.join(output_dir, arrear_name)

    arrear_free_df.to_excel(arrear_free_path, index=False)
    arrear_df.to_excel(arrear_path, index=False)

    format_excel_file(arrear_free_path)
    format_excel_file(arrear_path)

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
