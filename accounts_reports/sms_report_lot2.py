import os
import zipfile
import pandas as pd


LOOKUP_COLUMNS = [
    "cust_id",
    "status",
    "od_days",
    "total_arrear",
    "outstanding_principal",
    "outstanding_interest",
]


FILL_COLUMNS = [
    "status",
    "od_days",
    "total_arrear",
    "outstanding_principal",
    "outstanding_interest",
    "total_outstanding",
]


def norm_col(col):
    return str(col).strip().lower().replace(" ", "").replace("_", "")


def standardize_columns(df):
    rename_map = {}

    for col in df.columns:
        n = norm_col(col)

        if n == "custid":
            rename_map[col] = "cust_id"
        elif n == "status":
            rename_map[col] = "status"
        elif n == "oddays":
            rename_map[col] = "od_days"
        elif n == "totalarrear":
            rename_map[col] = "total_arrear"
        elif n == "outstandingprincipal":
            rename_map[col] = "outstanding_principal"
        elif n == "outstandinginterest":
            rename_map[col] = "outstanding_interest"
        elif n == "totaloutstanding":
            rename_map[col] = "total_outstanding"

    return df.rename(columns=rename_map)


def read_excel_clean(uploaded_file):
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]
    df = standardize_columns(df)
    return df


def build_writeoff_lookup(loan_os_file, loan_os_il_file):
    df1 = read_excel_clean(loan_os_file)
    df2 = read_excel_clean(loan_os_il_file)

    df = pd.concat([df1, df2], ignore_index=True)

    missing = [c for c in LOOKUP_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Loan OS Write Off files: {missing}")

    df = df[LOOKUP_COLUMNS].copy()

    df["status"] = df["status"].astype(str).str.strip()
    df["cust_id"] = df["cust_id"].astype(str).str.strip()

    # Keep only WriteOff cases
    df = df[df["status"].str.lower().eq("writeoff")].copy()

    for col in [
        "od_days",
        "total_arrear",
        "outstanding_principal",
        "outstanding_interest",
    ]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    lookup = (
        df.groupby("cust_id", as_index=False)
        .agg({
            "status": "first",
            "od_days": "max",
            "total_arrear": "sum",
            "outstanding_principal": "sum",
            "outstanding_interest": "sum",
        })
    )

    lookup["total_outstanding"] = (
        lookup["outstanding_principal"] +
        lookup["outstanding_interest"]
    )

    return lookup


def enrich_not_sent_writeoff_sms(not_sent_writeoff_file, lookup_df):
    base_df = read_excel_clean(not_sent_writeoff_file)

    if "cust_id" not in base_df.columns:
        raise ValueError("cust_id column not found in Not Sent Write Off SMS Data")

    base_df["cust_id"] = base_df["cust_id"].astype(str).str.strip()

    for col in FILL_COLUMNS:
        if col not in base_df.columns:
            base_df[col] = ""

    merged = base_df.merge(
        lookup_df,
        on="cust_id",
        how="left",
        suffixes=("", "_lookup")
    )

    for col in FILL_COLUMNS:
        lookup_col = f"{col}_lookup"

        if lookup_col in merged.columns:
            merged[col] = merged[lookup_col].combine_first(merged[col])
            merged.drop(columns=[lookup_col], inplace=True)

    return merged


def process_sms_report_lot2(files, output_dir, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)

    if progress_callback:
        progress_callback(10, "Reading Loan OS Write Off files...")

    lookup_df = build_writeoff_lookup(
        files["Loan OS Write Off"],
        files["Loan OS Write Off IL"]
    )

    if progress_callback:
        progress_callback(45, "Updating Not Sent Write Off SMS Data...")

    enriched_writeoff_df = enrich_not_sent_writeoff_sms(
        files["Not Sent Write Off SMS Data"],
        lookup_df
    )

    if progress_callback:
        progress_callback(70, "Reading Not Sent SMS Data...")

    not_sent_sms_df = read_excel_clean(files["Not Sent SMS Data"])

    not_sent_sms_path = os.path.join(output_dir, "Not Sent SMS Data.csv")
    not_sent_writeoff_path = os.path.join(
        output_dir,
        "Not Sent Write Off SMS Data Updated.csv"
    )
    lookup_path = os.path.join(
        output_dir,
        "Loan OS Write Off Consolidated Lookup.csv"
    )

    not_sent_sms_df.to_csv(
        not_sent_sms_path,
        index=False,
        encoding="utf-8-sig"
    )

    enriched_writeoff_df.to_csv(
        not_sent_writeoff_path,
        index=False,
        encoding="utf-8-sig"
    )

    lookup_df.to_csv(
        lookup_path,
        index=False,
        encoding="utf-8-sig"
    )

    if progress_callback:
        progress_callback(90, "Creating ZIP file...")

    zip_path = os.path.join(output_dir, "SMS_Report_Lot2_Output.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(not_sent_sms_path, arcname="Not Sent SMS Data.csv")
        zipf.write(
            not_sent_writeoff_path,
            arcname="Not Sent Write Off SMS Data Updated.csv"
        )
        zipf.write(
            lookup_path,
            arcname="Loan OS Write Off Consolidated Lookup.csv"
        )

    if progress_callback:
        progress_callback(100, "Completed.")

    summary = [
        {
            "report_name": "Not Sent SMS Data",
            "rows": len(not_sent_sms_df)
        },
        {
            "report_name": "Not Sent Write Off SMS Data Updated",
            "rows": len(enriched_writeoff_df)
        },
        {
            "report_name": "Loan OS Write Off Consolidated Lookup",
            "rows": len(lookup_df)
        }
    ]

    return summary, zip_path
