import os
import re
import tempfile
from datetime import datetime

import pandas as pd


MONTH_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]


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

    if name.endswith(".xlsb"):
        return "pyxlsb"

    return None


def next_month(month):
    return MONTH_ORDER[(MONTH_ORDER.index(month) + 1) % 12]


def prepare_repayment_summary(repayment_file):
    df = pd.read_excel(
        repayment_file,
        engine=get_engine(repayment_file),
        dtype=object
    )

    df = df.astype(object)

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

    return df.iloc[:last + 1].copy().astype(object)


def process_sheet(input_file, sheet_name, repayment_lookup):
    df = pd.read_excel(
        input_file,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl",
        dtype=object
    )

    df = df.astype(object)

    header_row, loan_col = find_header_and_loan_col(df)
    previous_closing_col, previous_month = find_previous_closing_col(df, header_row)

    df = trim_actual_rows(df, loan_col, header_row)

    current_month = next_month(previous_month)
    start_col = previous_closing_col + 1

    while df.shape[1] < start_col + 4:
        df[df.shape[1]] = ""

    df = df.astype(object)

    month_heading_row = header_row - 2
    total_row = header_row - 1

    df.iat[month_heading_row, start_col] = f"{current_month}-{datetime.now().strftime('%y')}"
    df.iat[header_row, start_col] = f"P.A {current_month}"
    df.iat[header_row, start_col + 1] = f"I.A {current_month}"
    df.iat[header_row, start_col + 2] = f"W.A {current_month}"
    df.iat[header_row, start_col + 3] = f"Closing {current_month}"

    total_pa = 0
    total_ia = 0
    total_wa = 0
    total_closing = 0

    for r in range(header_row + 1, len(df)):
        loan_id = df.iat[r, loan_col]

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

        previous_closing = to_number(df.iat[r, previous_closing_col])
        closing = previous_closing - pa

        df.iat[r, start_col] = display_value(pa)
        df.iat[r, start_col + 1] = display_value(ia)
        df.iat[r, start_col + 2] = display_value(wa)
        df.iat[r, start_col + 3] = display_value(closing)

        total_pa += pa
        total_ia += ia
        total_wa += wa
        total_closing += closing

    df.iat[total_row, start_col] = display_value(total_pa)
    df.iat[total_row, start_col + 1] = display_value(total_ia)
    df.iat[total_row, start_col + 2] = display_value(total_wa)
    df.iat[total_row, start_col + 3] = display_value(total_closing)

    return df, current_month


def write_df_to_sheet(writer, sheet_name, df):
    df = df.fillna("-")

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
        progress_callback(80, "Creating optimized output workbook...")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        write_df_to_sheet(writer, "System Write-Off", system_df)
        write_df_to_sheet(writer, "Manual Write-Off", manual_df)

    if progress_callback:
        progress_callback(100, "Report generated successfully.")

    return output_file
