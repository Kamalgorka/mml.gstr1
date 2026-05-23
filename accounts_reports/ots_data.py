import os
import tempfile
import pandas as pd


FINAL_COLUMNS = [
    "Hub", "Cluster ID", "Region", "Unit Name", "State", "District",
    "Branch_Id", "Branch_Name", "Finpage_Loan_No", "Cust_Id",
    "Funder_Description", "Member Name", "Mobile No.", "prod_category_id",
    "product_id", "Disbursement_Date", "Status", "Lo_Name", "center_name",
    "Principal_Arrear", "Interest_Arrear", "Total_Arrear",
    "Last_Maturity_Date", "Od_Days", "Outstanding_Principal",
    "Outstanding_Interest", "OTS Amount as on May 01",
]

COLUMN_MAP = {
    "Hub": ["zone", "ZONE"],
    "Cluster ID": ["cluster", "CLUSTER"],
    "Region": ["region", "REGION"],
    "Unit Name": ["unit", "UNIT"],
    "State": ["branch_state", "BRANCH STATE"],
    "District": ["BRANCH_DISTRICT", "BRANCH DISTRICT"],
    "Branch_Id": ["BranchCode", "branchcode"],
    "Branch_Name": ["branch_name"],
    "Finpage_Loan_No": ["finpage_loan_no"],
    "Cust_Id": ["cust_id"],
    "Funder_Description": ["Funder", "funder", "Funder_Description"],
    "Member Name": ["member_name"],
    "Mobile No.": ["mobile_number"],
    "prod_category_id": ["product_category_id"],
    "product_id": ["product_id"],
    "Disbursement_Date": ["disbursement_date", "disbursement_dateo"],
    "Status": ["status"],
    "Lo_Name": ["lo_name", "Loan Officer"],
    "center_name": ["center_name"],
    "Principal_Arrear": ["principal_arrear"],
    "Interest_Arrear": ["interest_arrear"],
    "Total_Arrear": ["total_arrear"],
    "Last_Maturity_Date": ["last_maturity_date"],
    "Od_Days": ["od_days"],
    "Outstanding_Principal": ["outstanding_principal"],
    "Outstanding_Interest": ["outstanding_interest"],
}

OUTSTANDING_ALLOWED_FUNDERS = {
    "ARCILARC_MARCH_2026",
    "CFMARC_MARCH_2025",
    "OWN",
    "PHOENIX_ARC",
    "PHOENIX_ARC-1",
    "BUSINESSLOAN_AL_AP",
    "PIRAMAL_DA_AUSTIN_09",
    "PIRAMAL_DA_ORCHID_05_2024",
    "",
}

PIRAMAL_60_DPD_FUNDERS = {
    "PIRAMAL_DA_AUSTIN_09",
    "PIRAMAL_DA_ORCHID_05_2024",
}


def _progress(progress_callback, value, message):
    if progress_callback:
        progress_callback(value, message)


def _norm_col(col):
    return str(col).strip().lower().replace(" ", "").replace("_", "").replace(".", "")


def _clean_text_series(series):
    return series.fillna("").astype(str).str.strip()


def _clean_numeric(series):
    return pd.to_numeric(
        series.fillna("").astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce"
    ).fillna(0)


def _normalize_status(series):
    return (
        _clean_text_series(series)
        .str.upper()
        .str.replace("-", " ", regex=False)
        .str.replace("_", " ", regex=False)
        .str.replace("  ", " ", regex=False)
        .str.strip()
    )


def _get_engine(filename):
    ext = os.path.splitext(filename.lower())[1]

    if ext == ".xlsb":
        return "pyxlsb"

    if ext == ".xls":
        return "xlrd"

    return None


def _read_uploaded_file(uploaded_file):
    filename = uploaded_file.name
    ext = os.path.splitext(filename.lower())[1]

    if ext == ".csv":
        return {"CSV": pd.read_csv(uploaded_file, dtype=str)}

    engine = _get_engine(filename)

    return pd.read_excel(
        uploaded_file,
        sheet_name=None,
        dtype=str,
        engine=engine
    )


def _combine_all_sheets(uploaded_file):
    sheets = _read_uploaded_file(uploaded_file)
    frames = []

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue

        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        df["Source_Sheet"] = sheet_name
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def _pick_column(df, possible_names):
    norm_lookup = {_norm_col(c): c for c in df.columns}

    for name in possible_names:
        key = _norm_col(name)

        if key in norm_lookup:
            return norm_lookup[key]

    return None


def _standardize_to_final(df, source_name):
    output = pd.DataFrame()

    for final_col in FINAL_COLUMNS:
        if final_col == "OTS Amount as on May 01":
            output[final_col] = ""
            continue

        source_col = _pick_column(df, COLUMN_MAP.get(final_col, []))

        if source_col:
            output[final_col] = df[source_col]
        else:
            output[final_col] = ""

    output["Source_File"] = source_name

    return output


def _filter_outstanding_data_with_cust_id_expansion(df):
    # Outstanding IL/JLG: only ACTIVE status rows
    status = _normalize_status(df["Status"])
    active_mask = status.isin(["ACTIVE"])

    df_active = df[active_mask].copy()

    # Base rows: funder allowed / blank / null
    funder = _clean_text_series(df_active["Funder_Description"]).str.upper()

    allowed_funder_mask = funder.isin(OUTSTANDING_ALLOWED_FUNDERS)
    blank_mask = funder.eq("") | funder.isin(["NAN", "NONE", "NULL"])

    base_mask = allowed_funder_mask | blank_mask

    # PIRAMAL funders only > 60 DPD
    od_days = _clean_numeric(df_active["Od_Days"])
    piramal_mask = funder.isin(PIRAMAL_60_DPD_FUNDERS)

    base_mask = base_mask & ((~piramal_mask) | (od_days > 60))

    base_df = df_active[base_mask].copy()

    # Take all ACTIVE rows of same Cust_Id also
    qualified_cust_ids = set(
        _clean_text_series(base_df["Cust_Id"])
        .replace("", pd.NA)
        .dropna()
        .astype(str)
    )

    cust_id_series = _clean_text_series(df_active["Cust_Id"]).astype(str)

    expanded_df = df_active[
        base_mask | cust_id_series.isin(qualified_cust_ids)
    ].copy()

    return expanded_df


def _filter_writeoff_data(df):
    # WriteOff IL/JLG: full cases where Status is WriteOff only
    status = _normalize_status(df["Status"])

    return df[
        status.isin(["WRITE OFF", "WRITEOFF"])
    ].copy()


def _calculate_ots_amount(df):
    principal = _clean_numeric(df["Outstanding_Principal"])
    interest = _clean_numeric(df["Outstanding_Interest"])

    df["OTS Amount as on May 01"] = principal + interest

    return df


def _optimize_output_types(df):
    numeric_cols = [
        "Principal_Arrear",
        "Interest_Arrear",
        "Total_Arrear",
        "Od_Days",
        "Outstanding_Principal",
        "Outstanding_Interest",
        "OTS Amount as on May 01",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = _clean_numeric(df[col])

    return df


def process_ots_data(
    outstanding_il_file,
    outstanding_jlg_file,
    writeoff_il_file,
    writeoff_jlg_file,
    progress_callback=None
):
    _progress(progress_callback, 5, "Reading Outstanding IL file...")

    outstanding_il = _standardize_to_final(
        _combine_all_sheets(outstanding_il_file),
        "Outstanding IL"
    )
    outstanding_il = _filter_outstanding_data_with_cust_id_expansion(outstanding_il)

    _progress(progress_callback, 25, "Reading Outstanding JLG file...")

    outstanding_jlg = _standardize_to_final(
        _combine_all_sheets(outstanding_jlg_file),
        "Outstanding JLG"
    )
    outstanding_jlg = _filter_outstanding_data_with_cust_id_expansion(outstanding_jlg)

    _progress(progress_callback, 45, "Reading Write Off IL file...")

    writeoff_il = _standardize_to_final(
        _combine_all_sheets(writeoff_il_file),
        "Write Off IL"
    )
    writeoff_il = _filter_writeoff_data(writeoff_il)

    _progress(progress_callback, 60, "Reading Write Off JLG file...")

    writeoff_jlg = _standardize_to_final(
        _combine_all_sheets(writeoff_jlg_file),
        "Write Off JLG"
    )
    writeoff_jlg = _filter_writeoff_data(writeoff_jlg)

    _progress(progress_callback, 75, "Combining final OTS data...")

    final_df = pd.concat(
        [outstanding_il, outstanding_jlg, writeoff_il, writeoff_jlg],
        ignore_index=True
    )

    final_df = _calculate_ots_amount(final_df)
    final_df = final_df[["Source_File"] + FINAL_COLUMNS]
    final_df = _optimize_output_types(final_df)

    _progress(progress_callback, 90, "Creating output file...")

    output_path = os.path.join(
        tempfile.gettempdir(),
        "OTS_Data_Output.xlsx"
    )

    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        engine_kwargs={
            "options": {
                "strings_to_numbers": True,
                "strings_to_urls": False,
                "constant_memory": True,
            }
        }
    ) as writer:
        final_df.to_excel(
            writer,
            index=False,
            sheet_name="OTS Data"
        )

        workbook = writer.book
        worksheet = writer.sheets["OTS Data"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "align": "center",
            "valign": "vcenter",
        })

        for col_num, value in enumerate(final_df.columns):
            worksheet.write(0, col_num, value, header_format)

        worksheet.freeze_panes(1, 0)

    _progress(progress_callback, 100, "OTS Data report generated successfully.")

    return output_path
