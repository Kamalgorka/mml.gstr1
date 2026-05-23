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
    return pd.read_excel(uploaded_file, sheet_name=None, dtype=str, engine=engine)


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

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


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
        output[final_col] = df[source_col] if source_col else ""

    output["Source_File"] = source_name
    return output


def _filter_outstanding_data(df):
    # Status: Active / WriteOff allowed
    status = _normalize_status(df["Status"])
    df = df[status.isin(["ACTIVE", "WRITE OFF", "WRITEOFF"])].copy()

    # Funder filter only for Outstanding IL/JLG
    funder = _clean_text_series(df["Funder_Description"]).str.upper()

    allowed_funder_mask = funder.isin(OUTSTANDING_ALLOWED_FUNDERS)
    blank_mask = funder.eq("") | funder.isin(["NAN", "NONE", "NULL"])

    df = df[allowed_funder_mask | blank_mask].copy()

    # Piramal cases only > 60 DPD
    funder_after = _clean_text_series(df["Funder_Description"]).str.upper()
    od_days = _clean_numeric(df["Od_Days"])

    piramal_mask = funder_after.isin(PIRAMAL_60_DPD_FUNDERS)
    df = df[(~piramal_mask) | (od_days > 60)].copy()

    return df


def _filter_writeoff_data(df):
    # WriteOff reports: no funder filter, only WriteOff status
    status = _normalize_status(df["Status"])
    return df[status.isin(["WRITE OFF", "WRITEOFF"])].copy()


def _calculate_ots_amount(df):
    principal = _clean_numeric(df["Outstanding_Principal"])
    interest = _clean_numeric(df["Outstanding_Interest"])
    df["OTS Amount as on May 01"] = principal + interest
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
    outstanding_il = _filter_outstanding_data(outstanding_il)

    _progress(progress_callback, 25, "Reading Outstanding JLG file...")
    outstanding_jlg = _standardize_to_final(
        _combine_all_sheets(outstanding_jlg_file),
        "Outstanding JLG"
    )
    outstanding_jlg = _filter_outstanding_data(outstanding_jlg)

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

    _progress(progress_callback, 90, "Creating output file...")

    output_path = os.path.join(tempfile.gettempdir(), "OTS_Data_Output.xlsx")

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, index=False, sheet_name="OTS Data")

        workbook = writer.book
        worksheet = writer.sheets["OTS Data"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })

        for col_num, value in enumerate(final_df.columns):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 18)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(final_df), len(final_df.columns) - 1)

    _progress(progress_callback, 100, "OTS Data report generated successfully.")
    return output_path
