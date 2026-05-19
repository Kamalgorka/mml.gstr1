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
        elif n == "maxoddays":
            rename_map[col] = "max_od_days"
        elif n == "totalarrear":
            rename_map[col] = "total_arrear"
        elif n == "totalarrearsum":
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


def normalize_cust_id(df):
    if "cust_id" in df.columns:
        df["cust_id"] = (
            df["cust_id"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )
    return df


def add_discrepancy_column(df, check_total_arrear=False):
    required_cols = [
        "outstanding_principal",
        "outstanding_interest",
        "total_outstanding",
    ]

    if check_total_arrear:
        required_cols.append("total_arrear")

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns for validation: {missing}")

    for col in required_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    calculated_total = df["outstanding_principal"] + df["outstanding_interest"]

    mismatch_mask = calculated_total.round(2) != df["total_outstanding"].round(2)
    total_zero_negative_mask = df["total_outstanding"] <= 0
    principal_negative_mask = df["outstanding_principal"] < 0
    interest_negative_mask = df["outstanding_interest"] < 0

    discrepancy = pd.Series("", index=df.index, dtype="object")

    discrepancy.loc[mismatch_mask] += (
        "total_outstanding not equal to outstanding_principal + outstanding_interest; "
    )

    discrepancy.loc[total_zero_negative_mask] += (
        "total_outstanding is zero/negative; "
    )

    discrepancy.loc[principal_negative_mask] += (
        "outstanding_principal is negative; "
    )

    discrepancy.loc[interest_negative_mask] += (
        "outstanding_interest is negative; "
    )

    if check_total_arrear:
        total_arrear_negative_mask = df["total_arrear"] < 0
        outstanding_less_than_arrear_mask = df["total_outstanding"] < df["total_arrear"]

        discrepancy.loc[total_arrear_negative_mask] += (
            "total_arrear is negative; "
        )

        discrepancy.loc[outstanding_less_than_arrear_mask] += (
            "total_outstanding is less than total_arrear; "
        )

    df["Discrepancy"] = discrepancy.str.strip("; ")

    return df


def consolidate_arrear_free_into_arrear(arrear_free_df, arrear_df):
    """
    If same cust_id exists in both:
    - Add Arrear Free amount values into Arrear same cust_id row
    - max_od_days = max of Arrear / Arrear Free
    - Remove matched cust_id rows from Arrear Free output
    No extra output file is created.
    """
    if arrear_free_df.empty or arrear_df.empty:
        return arrear_free_df, arrear_df

    arrear_free_df = normalize_cust_id(arrear_free_df.copy())
    arrear_df = normalize_cust_id(arrear_df.copy())

    required_cols = [
        "cust_id",
        "max_od_days",
        "total_arrear",
        "outstanding_principal",
        "outstanding_interest",
        "total_outstanding",
    ]

    missing_free = [c for c in required_cols if c not in arrear_free_df.columns]
    missing_arrear = [c for c in required_cols if c not in arrear_df.columns]

    if missing_free:
        raise ValueError(f"Missing columns in Arrear Free data: {missing_free}")

    if missing_arrear:
        raise ValueError(f"Missing columns in Arrear data: {missing_arrear}")

    matched_ids = set(arrear_free_df["cust_id"]).intersection(
        set(arrear_df["cust_id"])
    )

    if not matched_ids:
        return arrear_free_df, arrear_df

    numeric_cols = [
        "max_od_days",
        "total_arrear",
        "outstanding_principal",
        "outstanding_interest",
        "total_outstanding",
    ]

    for col in numeric_cols:
        arrear_free_df[col] = pd.to_numeric(arrear_free_df[col], errors="coerce").fillna(0)
        arrear_df[col] = pd.to_numeric(arrear_df[col], errors="coerce").fillna(0)

    free_matched = arrear_free_df[
        arrear_free_df["cust_id"].isin(matched_ids)
    ].copy()

    free_grouped = (
        free_matched
        .groupby("cust_id", as_index=False)
        .agg({
            "max_od_days": "max",
            "total_arrear": "sum",
            "outstanding_principal": "sum",
            "outstanding_interest": "sum",
            "total_outstanding": "sum",
        })
    )

    arrear_df = arrear_df.merge(
        free_grouped,
        on="cust_id",
        how="left",
        suffixes=("", "_free")
    )

    for col in [
        "total_arrear",
        "outstanding_principal",
        "outstanding_interest",
        "total_outstanding",
    ]:
        arrear_df[col] = arrear_df[col] + arrear_df[f"{col}_free"].fillna(0)
        arrear_df.drop(columns=[f"{col}_free"], inplace=True)

    arrear_df["max_od_days"] = arrear_df[
        ["max_od_days", "max_od_days_free"]
    ].max(axis=1)

    arrear_df.drop(columns=["max_od_days_free"], inplace=True)

    arrear_free_df = arrear_free_df[
        ~arrear_free_df["cust_id"].isin(matched_ids)
    ].copy()

    # Recalculate discrepancies after consolidation
    arrear_free_df = add_discrepancy_column(
        arrear_free_df,
        check_total_arrear=False
    )

    arrear_df = add_discrepancy_column(
        arrear_df,
        check_total_arrear=True
    )

    return arrear_free_df, arrear_df


def build_writeoff_lookup(loan_os_file, loan_os_il_file):
    df1 = read_excel_clean(loan_os_file)
    df2 = read_excel_clean(loan_os_il_file)

    df = pd.concat([df1, df2], ignore_index=True)
    df = normalize_cust_id(df)

    missing = [c for c in LOOKUP_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Loan OS Write Off files: {missing}")

    df = df[LOOKUP_COLUMNS].copy()

    df["status"] = df["status"].astype(str).str.strip()

    # Keep only WriteOff cases before consolidation
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
    base_df = normalize_cust_id(base_df)

    if "cust_id" not in base_df.columns:
        raise ValueError("cust_id column not found in Not Sent Write Off SMS Data")

    lookup_df = normalize_cust_id(lookup_df)

    for col in FILL_COLUMNS:
        if col not in base_df.columns:
            base_df[col] = ""

    # Inner join removes rows where cust_id is not available after WriteOff-only filter
    merged = base_df.merge(
        lookup_df,
        on="cust_id",
        how="inner",
        suffixes=("", "_lookup")
    )

    for col in FILL_COLUMNS:
        lookup_col = f"{col}_lookup"

        if lookup_col in merged.columns:
            merged[col] = merged[lookup_col].combine_first(merged[col])
            merged.drop(columns=[lookup_col], inplace=True)

    merged = add_discrepancy_column(
        merged,
        check_total_arrear=True
    )

    return merged


def process_not_sent_sms_data(not_sent_sms_file):
    df = read_excel_clean(not_sent_sms_file)
    df = normalize_cust_id(df)

    if "status" not in df.columns:
        raise ValueError("status column not found in Not Sent SMS Data")

    if "max_od_days" not in df.columns:
        raise ValueError("max_od_days column not found in Not Sent SMS Data")

    df["status"] = df["status"].astype(str).str.strip()

    # Keep only Active cases
    df = df[df["status"].str.lower().eq("active")].copy()

    df["max_od_days"] = pd.to_numeric(df["max_od_days"], errors="coerce").fillna(0)

    arrear_free_df = df[df["max_od_days"] <= 0].copy()
    arrear_df = df[df["max_od_days"] > 0].copy()

    arrear_free_df = add_discrepancy_column(
        arrear_free_df,
        check_total_arrear=False
    )

    arrear_df = add_discrepancy_column(
        arrear_df,
        check_total_arrear=True
    )

    arrear_free_df, arrear_df = consolidate_arrear_free_into_arrear(
        arrear_free_df,
        arrear_df
    )

    return arrear_free_df, arrear_df, len(df)


def process_sms_report_lot2(files, output_dir, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)

    if progress_callback:
        progress_callback(10, "Reading Loan OS Write Off files...")

    lookup_df = build_writeoff_lookup(
        files["Loan OS Write Off"],
        files["Loan OS Write Off IL"]
    )

    if progress_callback:
        progress_callback(35, "Updating and validating Not Sent Write Off SMS Data...")

    enriched_writeoff_df = enrich_not_sent_writeoff_sms(
        files["Not Sent Write Off SMS Data"],
        lookup_df
    )

    if progress_callback:
        progress_callback(60, "Processing Not Sent SMS Data...")

    arrear_free_df, arrear_df, active_rows = process_not_sent_sms_data(
        files["Not Sent SMS Data"]
    )

    not_sent_arrear_free_path = os.path.join(
        output_dir,
        "Arrear Free Not Sent SMS Data.csv"
    )

    not_sent_arrear_path = os.path.join(
        output_dir,
        "Arrear Not Sent SMS Data.csv"
    )

    not_sent_writeoff_path = os.path.join(
        output_dir,
        "Not Sent Write Off SMS Data Updated.csv"
    )

    lookup_path = os.path.join(
        output_dir,
        "Loan OS Write Off Consolidated Lookup.csv"
    )

    arrear_free_df.to_csv(
        not_sent_arrear_free_path,
        index=False,
        encoding="utf-8-sig"
    )

    arrear_df.to_csv(
        not_sent_arrear_path,
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
        zipf.write(
            not_sent_arrear_free_path,
            arcname="Arrear Free Not Sent SMS Data.csv"
        )

        zipf.write(
            not_sent_arrear_path,
            arcname="Arrear Not Sent SMS Data.csv"
        )

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
            "report_name": "Arrear Free Not Sent SMS Data",
            "rows": len(arrear_free_df),
            "discrepancy_rows": int((arrear_free_df["Discrepancy"] != "").sum())
        },
        {
            "report_name": "Arrear Not Sent SMS Data",
            "rows": len(arrear_df),
            "discrepancy_rows": int((arrear_df["Discrepancy"] != "").sum())
        },
        {
            "report_name": "Not Sent Write Off SMS Data Updated",
            "rows": len(enriched_writeoff_df),
            "discrepancy_rows": int((enriched_writeoff_df["Discrepancy"] != "").sum())
        },
        {
            "report_name": "Loan OS Write Off Consolidated Lookup",
            "rows": len(lookup_df)
        },
        {
            "report_name": "Active cases considered from Not Sent SMS Data",
            "rows": active_rows
        }
    ]

    return summary, zip_path
