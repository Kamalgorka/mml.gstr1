import os
import tempfile
import pandas as pd


EXCLUDE_FUNDERS = {
    "ARCILARC_MARCH_2026",
    "CFMARC_MARCH_2025",
    "PHOENIX_ARC",
    "PHOENIX_ARC-1",
}


def norm_col(col):
    return str(col).strip().lower().replace(" ", "_")


def find_col(df, possible_names):
    col_map = {norm_col(c): c for c in df.columns}

    for name in possible_names:
        key = norm_col(name)
        if key in col_map:
            return col_map[key]

    raise ValueError(f"Required column not found. Expected one of: {possible_names}")


def read_uploaded_file(uploaded_file):
    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    if filename.endswith(".xlsb"):
        return pd.read_excel(uploaded_file, engine="pyxlsb")

    return pd.read_excel(uploaded_file)


def process_single_client_od(uploaded_file):
    df = read_uploaded_file(uploaded_file)

    # Required columns with flexible naming
    status_col = find_col(df, ["status", "Status"])
    funder_col = find_col(df, ["funder", "Funder", "Funder_Description", "funder_description"])
    od_days_col = find_col(df, ["od_days", "OD Days", "Od_Days", "OD_Days"])
    center_col = find_col(df, ["center_name", "Center Name", "CENTER_NAME"])
    cust_col = find_col(df, ["cust_id", "Cust ID", "Cust_Id", "customer_id"])

    # Clean working values
    df[status_col] = df[status_col].astype(str).str.strip()
    df[funder_col] = df[funder_col].astype(str).str.strip()
    df[center_col] = df[center_col].astype(str).str.strip()
    df[cust_col] = df[cust_col].astype(str).str.strip()

    df[od_days_col] = pd.to_numeric(df[od_days_col], errors="coerce")

    # 1. Remove Death cases from Status
    df = df[~df[status_col].str.upper().eq("DEATH")]

    # 2. Remove excluded funders
    df = df[~df[funder_col].str.upper().isin({x.upper() for x in EXCLUDE_FUNDERS})]

    # 3. Keep OD days <= 90
    df = df[df[od_days_col] <= 90]

    # Remove blank center/cust_id before grouping
    df = df[
        df[center_col].notna()
        & df[cust_col].notna()
        & (df[center_col].astype(str).str.strip() != "")
        & (df[cust_col].astype(str).str.strip() != "")
        & (df[cust_col].astype(str).str.lower().str.strip() != "nan")
    ]

    # 4. Keep only centers where unique cust_id count = 1
    center_unique_count = df.groupby(center_col)[cust_col].nunique(dropna=True)

    single_client_centers = center_unique_count[
        center_unique_count == 1
    ].index

    final_df = df[df[center_col].isin(single_client_centers)].copy()

    # Save output
    output_dir = tempfile.mkdtemp()
    output_path = os.path.join(output_dir, "Single_Client_OD_in_Center.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Single Client OD", index=False)

    return output_path
