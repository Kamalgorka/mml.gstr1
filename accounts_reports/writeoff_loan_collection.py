import os
import re
import tempfile
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


MONTH_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]


def clean_header(value):
    if value is None:
        return ""
    return str(value).strip().replace("\n", " ").replace("\r", " ")


def norm_col(value):
    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace(".", "")
    )


def to_number(value):
    if value is None or value == "":
        return 0

    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if value in ["", "-", "nan", "NaN"]:
            return 0

    try:
        return float(value)
    except Exception:
        return 0


def display_value(value):
    value = round(float(value), 0)

    if value == 0:
        return "-"

    return int(value)


def find_header_row_and_loan_col(ws):
    for row in range(1, min(ws.max_row, 15) + 1):
        for col in range(1, ws.max_column + 1):
            val = clean_header(ws.cell(row=row, column=col).value)
            if norm_col(val) == "loanid":
                return row, col

    raise Exception(f"Loan ID column not found in sheet: {ws.title}")


def find_last_month_block(ws, header_row):
    month_cols = []

    for col in range(1, ws.max_column + 1):
        val = clean_header(ws.cell(row=header_row, column=col).value)

        match = re.match(r"Closing\s+([A-Za-z]{3,})", val, re.IGNORECASE)
        if match:
            month_name = match.group(1).[:3].title()
            if month_name in MONTH_ORDER:
                month_cols.append((col, month_name))

    if not month_cols:
        raise Exception(f"No Closing month column found in sheet: {ws.title}")

    last_col, last_month = month_cols[-1]
    return last_col, last_month


def next_month_name(month_name):
    idx = MONTH_ORDER.index(month_name)
    return MONTH_ORDER[(idx + 1) % 12]


def copy_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell._style = src_cell._style.copy()

    if src_cell.number_format:
        dst_cell.number_format = src_cell.number_format

    if src_cell.alignment:
        dst_cell.alignment = src_cell.alignment.copy()


def prepare_repayment_summary(repayment_file):
    df = pd.read_excel(repayment_file)

    rename_map = {}
    for col in df.columns:
        n = norm_col(col)

        if n == "loanid":
            rename_map[col] = "loan_id"
        elif n == "principalcollected":
            rename_map[col] = "principal_collected"
        elif n == "interestcollected":
            rename_map[col] = "interest_collected"
        elif n in ["waiveoffamt", "waiveoffamount", "waiveoff"]:
            rename_map[col] = "waiveoff_amt"

    df = df.rename(columns=rename_map)

    required = [
        "loan_id",
        "principal_collected",
        "interest_collected",
        "waiveoff_amt",
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise Exception(f"Missing columns in repayment file: {missing}")

    df["loan_id"] = df["loan_id"].astype(str).str.strip()

    for col in ["principal_collected", "interest_collected", "waiveoff_amt"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    summary = (
        df.groupby("loan_id", as_index=False)
        .agg(
            principal_collected=("principal_collected", "sum"),
            interest_collected=("interest_collected", "sum"),
            waiveoff_amt=("waiveoff_amt", "sum"),
        )
    )

    return summary.set_index("loan_id").to_dict("index")


def update_writeoff_sheet(ws, repayment_lookup):
    header_row, loan_col = find_header_row_and_loan_col(ws)

    previous_closing_col, previous_month = find_last_month_block(ws, header_row)
    current_month = next_month_name(previous_month)

    start_col = previous_closing_col + 1

    headers = [
        f"P.A {current_month}",
        f"I.A {current_month}",
        f"W.A {current_month}",
        f"Closing {current_month}",
    ]

    # Month heading row
    month_heading_row = header_row - 2 if header_row > 2 else 1
    amount_row = header_row - 1 if header_row > 1 else header_row

    ws.cell(row=month_heading_row, column=start_col).value = current_month + "-" + datetime.now().strftime("%y")
    ws.merge_cells(
        start_row=month_heading_row,
        start_column=start_col,
        end_row=month_heading_row,
        end_column=start_col + 3,
    )

    month_cell = ws.cell(row=month_heading_row, column=start_col)
    month_cell.font = Font(bold=True)
    month_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Column headers
    for i, header in enumerate(headers):
        cell = ws.cell(row=header_row, column=start_col + i)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

        # Copy width/style from previous month block
        old_col = previous_closing_col - 3 + i
        copy_style(ws.cell(row=header_row, column=old_col), cell)
        ws.column_dimensions[get_column_letter(start_col + i)].width = ws.column_dimensions[
            get_column_letter(old_col)
        ].width

    # Fill data
    total_pa = 0
    total_ia = 0
    total_wa = 0
    total_closing = 0

    for row in range(header_row + 1, ws.max_row + 1):
        loan_id = ws.cell(row=row, column=loan_col).value

        if loan_id is None or str(loan_id).strip() == "":
            continue

        loan_id = str(loan_id).strip()

        data = repayment_lookup.get(
            loan_id,
            {
                "principal_collected": 0,
                "interest_collected": 0,
                "waiveoff_amt": 0,
            },
        )

        pa = data["principal_collected"]
        ia = data["interest_collected"]
        wa = data["waiveoff_amt"]

        previous_closing = to_number(ws.cell(row=row, column=previous_closing_col).value)
        closing = previous_closing - pa

        values = [
            display_value(pa),
            display_value(ia),
            display_value(wa),
            display_value(closing),
        ]

        for i, val in enumerate(values):
            cell = ws.cell(row=row, column=start_col + i)
            cell.value = val
            cell.alignment = Alignment(horizontal="center", vertical="center")

            old_col = previous_closing_col - 3 + i
            copy_style(ws.cell(row=row, column=old_col), cell)

        total_pa += pa
        total_ia += ia
        total_wa += wa
        total_closing += closing

    # Total row values
    totals = [
        display_value(total_pa),
        display_value(total_ia),
        display_value(total_wa),
        display_value(total_closing),
    ]

    for i, val in enumerate(totals):
        cell = ws.cell(row=amount_row, column=start_col + i)
        cell.value = val
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    return current_month


def process_writeoff_loan_collection(writeoff_file, repayment_file, output_dir=None, progress_callback=None):
    if progress_callback:
        progress_callback(10, "Reading repayment file...")

    repayment_lookup = prepare_repayment_summary(repayment_file)

    if progress_callback:
        progress_callback(30, "Opening WriteOff loan collection file...")

    wb = load_workbook(writeoff_file)

    required_sheets = ["System Write-Off", "Manual Write-Off"]
    missing_sheets = [s for s in required_sheets if s not in wb.sheetnames]

    if missing_sheets:
        raise Exception(f"Missing sheets in WriteOff file: {missing_sheets}")

    if progress_callback:
        progress_callback(50, "Updating System Write-Off sheet...")

    system_month = update_writeoff_sheet(wb["System Write-Off"], repayment_lookup)

    if progress_callback:
        progress_callback(75, "Updating Manual Write-Off sheet...")

    manual_month = update_writeoff_sheet(wb["Manual Write-Off"], repayment_lookup)

    if output_dir is None:
        output_dir = tempfile.gettempdir()

    output_file = os.path.join(
        output_dir,
        f"WriteOff_Loan_Collection_Updated_{system_month}.xlsx",
    )

    wb.save(output_file)

    if progress_callback:
        progress_callback(100, "Report generated successfully.")

    return output_file
