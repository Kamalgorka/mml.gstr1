import os
import re
import tempfile
from datetime import datetime

import pandas as pd


MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

BASE_OUTPUT_COLS = 10  # A:J only


def norm_col(v):
    return str(v).strip().lower().replace(" ", "").replace("_", "").replace(".", "")


def to_number(v):
    if pd.isna(v):
        return 0
    if str(v).strip() in ["", "-"]:
        return 0
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0


def display_value(v):
    v = round(float(v), 0)
    return "-" if v == 0 else int(v)


def get_engine(file):
    name = getattr(file, "name", str(file)).lower()
    return "pyxlsb" if name.endswith(".xlsb") else None


def next_month(month):
    return MONTH_ORDER[(MONTH_ORDER.index(month) + 1) % 12]


def prepare_repayment_summary(repayment_file):
    df = pd.read_excel(
        repayment_file,
        engine=get_engine(repayment_file),
        dtype=object
    )

    rename = {}

    for col in df.columns:
        n = norm_col(col)

        if n == "loanid":
            rename[col] = "loan_id"
        elif n == "principalcollected":
            rename[col] = "principal_collected"
        elif n == "interestcollected":
            rename[col] = "interest_collected"
        elif n in ["waiveoffamt", "waiveoffamount", "waiveoff"]:
            rename[col] = "waiveoff_amt"

    df = df.rename(columns=rename)

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

    for c in ["principal_collected", "interest_collected", "waiveoff_amt"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    summary = (
        df.groupby("loan_id", as_index=False)
        .agg(
            principal_collected=("principal_collected", "sum"),
            interest_collected=("interest_collected", "sum"),
            waiveoff_amt=("waiveoff_amt", "sum"),
        )
    )

    return summary.set_index("loan_id").to_dict("index")


def find_header_and_loan_col(df):
    for r in range(min(15, len(df))):
        for c in range(len(df.columns)):
            if norm_col(df.iat[r, c]) == "loanid":
                return r, c

    raise Exception("Loan ID column not found.")


def find_previous_closing_col(df, header_row):
    closing_cols = []

    for c in range(len(df.columns)):
        val = str(df.iat[header_row, c]).strip()
        match = re.match(r"Closing\s+([A-Za-z]{3,})", val, re.IGNORECASE)

        if match:
            month = match.group(1)[:3].title()
            if month in MONTH_ORDER:
                closing_cols.append((c, month))

    if not closing_cols:
        raise Exception("Previous Closing month column not found.")

    return closing_cols[-1]


def trim_actual_rows(df, loan_col, header_row):
    last = header_row
    blank_count = 0

    for r in range(header_row + 1, len(df)):
        val = df.iat[r, loan_col]

        if not pd.isna(val) and str(val).strip() != "":
            last = r
            blank_count = 0
        else:
            blank_count += 1

        if blank_count >= 200 and last > header_row:
            break

    return df.iloc[:last + 1].copy()


def process_sheet(input_file, sheet_name, repayment_lookup):
    full_df = pd.read_excel(
        input_file,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
        dtype=object
    )

    header_row, loan_col = find_header_and_loan_col(full_df)
    previous_closing_col, previous_month = find_previous_closing_col(full_df, header_row)

    full_df = trim_actual_rows(full_df, loan_col, header_row)

    current_month = next_month(previous_month)

    # Output will contain only A:J from original sheet
    output_df = full_df.iloc[:, :BASE_OUTPUT_COLS].copy().astype(object)

    pa_col = BASE_OUTPUT_COLS
    ia_col = BASE_OUTPUT_COLS + 1
    wa_col = BASE_OUTPUT_COLS + 2
    closing_col = BASE_OUTPUT_COLS + 3

    for col in [pa_col, ia_col, wa_col, closing_col]:
        output_df[col] = pd.Series([""] * len(output_df), dtype=object)

    output_df = output_df.astype(object)

    month_heading_row = header_row - 2
    total_row = header_row - 1

    output_df.iat[month_heading_row, pa_col] = f"{current_month}-{datetime.now().strftime('%y')}"

    output_df.iat[header_row, pa_col] = f"P.A {current_month}"
    output_df.iat[header_row, ia_col] = f"I.A {current_month}"
    output_df.iat[header_row, wa_col] = f"W.A {current_month}"
    output_df.iat[header_row, closing_col] = f"Closing {current_month}"

    total_pa = 0
    total_ia = 0
    total_wa = 0
    total_closing = 0

    for r in range(header_row + 1, len(full_df)):
        loan_id = full_df.iat[r, loan_col]

        if pd.isna(loan_id) or str(loan_id).strip() == "":
            continue

        loan_id = str(loan_id).strip()

        data = repayment_lookup.get(
            loan_id,
            {
                "principal_collected": 0,
                "interest_collected": 0,
                "waiveoff_amt": 0,
            }
        )

        pa = data["principal_collected"]
        ia = data["interest_collected"]
        wa = data["waiveoff_amt"]

        previous_closing = to_number(full_df.iat[r, previous_closing_col])
        closing = previous_closing - pa

        output_df.iat[r, pa_col] = display_value(pa)
        output_df.iat[r, ia_col] = display_value(ia)
        output_df.iat[r, wa_col] = display_value(wa)
        output_df.iat[r, closing_col] = display_value(closing)

        total_pa += pa
        total_ia += ia
        total_wa += wa
        total_closing += closing

    output_df.iat[total_row, pa_col] = display_value(total_pa)
    output_df.iat[total_row, ia_col] = display_value(total_ia)
    output_df.iat[total_row, wa_col] = display_value(total_wa)
    output_df.iat[total_row, closing_col] = display_value(total_closing)

    output_df = output_df.fillna("-")

    return output_df, current_month


def write_df_to_sheet(writer, sheet_name, df):
    df.to_excel(
        writer,
        sheet_name=sheet_name,
        index=False,
        header=False
    )


def process_writeoff_loan_collection(
    writeoff_file,
    repayment_file,
    output_dir=None,
    progress_callback=None
):
    if progress_callback:
        progress_callback(10, "Reading repayment file...")

    repayment_lookup = prepare_repayment_summary(repayment_file)

    if progress_callback:
        progress_callback(30, "Processing System Write-Off sheet...")

    system_df, system_month = process_sheet(
        writeoff_file,
        "System Write-Off",
        repayment_lookup
    )

    if progress_callback:
        progress_callback(60, "Processing Manual Write-Off sheet...")

    manual_df, manual_month = process_sheet(
        writeoff_file,
        "Manual Write-Off",
        repayment_lookup
    )

    if output_dir is None:
        output_dir = tempfile.gettempdir()

    output_file = os.path.join(
        output_dir,
        f"WriteOff_Loan_Collection_Updated_{system_month}.xlsx"
    )

    if progress_callback:
        progress_callback(85, "Creating final output workbook...")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        write_df_to_sheet(writer, "System Write-Off", system_df)
        write_df_to_sheet(writer, "Manual Write-Off", manual_df)

    if progress_callback:
        progress_callback(100, "Report generated successfully.")

    return output_file
