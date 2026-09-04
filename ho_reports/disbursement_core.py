import os
from datetime import datetime, timedelta

import numpy as np
import pandas as pd


# ============================================================
# REQUIRED SOURCE FILES
# ============================================================

REQUIRED_FILES = {
    "ICICI.xlsx": "icici_jlg",
    "ICICI IL.xlsx": "icici_il",
    "CASHLESS DISBURSEMENT.xlsx": "db_jlg",
    "CASHLESS DISBURSEMENT IL.xlsx": "db_il",
    "MASTER DATA NEFT.xlsx": "master",
    "ARC.xlsx": "arc",
}


# ============================================================
# FILE IDENTIFICATION
# ============================================================

def clean_uploaded_filename(name):
    """
    Folder upload can sometimes provide a relative path such as:

        Disb Data/ICICI.xlsx

    We only need the actual filename.
    """

    name = str(name).replace("\\", "/")
    return name.split("/")[-1].strip()


def normalize_filename(name):
    return clean_uploaded_filename(name).lower()


def identify_uploaded_files(uploaded_files):
    """
    Automatically identifies the six required files
    from the folder selected by the user.
    """

    uploaded_map = {}

    for uploaded_file in uploaded_files:
        filename = normalize_filename(uploaded_file.name)

        # If same filename occurs more than once,
        # retain first occurrence.
        if filename not in uploaded_map:
            uploaded_map[filename] = uploaded_file

    identified = {}
    missing = []

    for required_filename, internal_key in REQUIRED_FILES.items():

        expected = required_filename.lower()

        if expected in uploaded_map:
            identified[internal_key] = uploaded_map[expected]
        else:
            missing.append(required_filename)

    return identified, missing


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text_series(series):
    """
    Convert values to cleaned strings without destroying blanks.
    """

    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
    )


def clean_id_series(series):
    """
    Normalize IDs that Excel may sometimes read as:

        12345
        12345.0

    This assists matching while retaining text IDs.
    """

    cleaned = clean_text_series(series)

    cleaned = cleaned.str.replace(
        r"\.0$",
        "",
        regex=True
    )

    return cleaned


def clean_account_series(series):
    return clean_text_series(series)


def clean_ifsc_series(series):
    return (
        clean_text_series(series)
        .str.upper()
    )


def validate_columns(df, required_columns, file_name):
    """
    Stop calculation if source structure is incorrect.
    """

    missing = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing:
        raise ValueError(
            f"{file_name} is missing required column(s): "
            + ", ".join(missing)
        )


def parse_date_column(series):
    """
    Flexible conversion because Posting-Date may come
    in different Excel/date formats.
    """

    return pd.to_datetime(
        series,
        errors="coerce",
        dayfirst=True
    )


def get_product_column(df):
    """
    After merge, Product may appear as:

        Product
        Product_x
        Product_y

    depending upon the columns available in source files.
    """

    possible_columns = [
        "Product",
        "Product_x",
        "Product_y",
    ]

    for column in possible_columns:
        if column in df.columns:
            return column

    return None


# ============================================================
# MAIN CALCULATION ENGINE
# ============================================================

def run_reports(files):

    # ========================================================
    # READ SOURCE FILES
    # ========================================================

    icici_sheet = pd.read_excel(
        files["icici_jlg"],
        engine="openpyxl",
        dtype=str
    )

    icici_sheet_il = pd.read_excel(
        files["icici_il"],
        engine="openpyxl",
        dtype=str
    )

    master_sheet = pd.read_excel(
        files["master"],
        engine="openpyxl",
        dtype=str
    )

    disbursement_sheet = pd.read_excel(
        files["db_jlg"],
        engine="openpyxl",
        dtype=str
    )

    disbursement_sheet_il = pd.read_excel(
        files["db_il"],
        engine="openpyxl",
        dtype=str
    )

    arc_sheet = pd.read_excel(
        files["arc"],
        engine="openpyxl",
        dtype=str
    )

    # ========================================================
    # CLEAN COLUMN HEADINGS
    # ========================================================

    all_dataframes = [
        icici_sheet,
        icici_sheet_il,
        master_sheet,
        disbursement_sheet,
        disbursement_sheet_il,
        arc_sheet,
    ]

    for df in all_dataframes:
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

    # ========================================================
    # REQUIRED COLUMN VALIDATION
    # ========================================================

    validate_columns(
        icici_sheet,
        [
            "Coustmer ID",
            "A/C no.",
            "Loan Account ID",
            "IFSC Code",
            "Amount",
        ],
        "ICICI.xlsx"
    )

    validate_columns(
        icici_sheet_il,
        [
            "Coustmer ID",
            "A/C no.",
            "Loan Account ID",
            "IFSC Code",
            "Amount",
        ],
        "ICICI IL.xlsx"
    )

    validate_columns(
        master_sheet,
        [
            "Coustmer ID",
            "A/C no.",
            "Loan Account ID",
            "Posting-Date",
            "IFSC Code",
        ],
        "MASTER DATA NEFT.xlsx"
    )

    validate_columns(
        disbursement_sheet,
        [
            "Loan ID",
            "IFSC Code",
            "Customer ID",
            "Final Disbursal Amount",
            "Account",
        ],
        "CASHLESS DISBURSEMENT.xlsx"
    )

    validate_columns(
        disbursement_sheet_il,
        [
            "Loan ID",
            "IFSC Code",
            "Customer ID",
            "Final Disbursal Amount",
            "Account",
        ],
        "CASHLESS DISBURSEMENT IL.xlsx"
    )

    validate_columns(
        arc_sheet,
        [
            "Member ID"
        ],
        "ARC.xlsx"
    )

    # ========================================================
    # CLEAN MATCHING FIELDS
    # ========================================================

    # ICICI JLG
    icici_sheet["Coustmer ID"] = clean_id_series(
        icici_sheet["Coustmer ID"]
    )

    icici_sheet["Loan Account ID"] = clean_id_series(
        icici_sheet["Loan Account ID"]
    )

    icici_sheet["A/C no."] = clean_account_series(
        icici_sheet["A/C no."]
    )

    icici_sheet["IFSC Code"] = clean_ifsc_series(
        icici_sheet["IFSC Code"]
    )

    # ICICI IL
    icici_sheet_il["Coustmer ID"] = clean_id_series(
        icici_sheet_il["Coustmer ID"]
    )

    icici_sheet_il["Loan Account ID"] = clean_id_series(
        icici_sheet_il["Loan Account ID"]
    )

    icici_sheet_il["A/C no."] = clean_account_series(
        icici_sheet_il["A/C no."]
    )

    icici_sheet_il["IFSC Code"] = clean_ifsc_series(
        icici_sheet_il["IFSC Code"]
    )

    # MASTER
    master_sheet["Coustmer ID"] = clean_id_series(
        master_sheet["Coustmer ID"]
    )

    master_sheet["Loan Account ID"] = clean_id_series(
        master_sheet["Loan Account ID"]
    )

    master_sheet["A/C no."] = clean_account_series(
        master_sheet["A/C no."]
    )

    master_sheet["IFSC Code"] = clean_ifsc_series(
        master_sheet["IFSC Code"]
    )

    # DB JLG
    disbursement_sheet["Loan ID"] = clean_id_series(
        disbursement_sheet["Loan ID"]
    )

    disbursement_sheet["Customer ID"] = clean_id_series(
        disbursement_sheet["Customer ID"]
    )

    disbursement_sheet["Account"] = clean_account_series(
        disbursement_sheet["Account"]
    )

    disbursement_sheet["IFSC Code"] = clean_ifsc_series(
        disbursement_sheet["IFSC Code"]
    )

    # DB IL
    disbursement_sheet_il["Loan ID"] = clean_id_series(
        disbursement_sheet_il["Loan ID"]
    )

    disbursement_sheet_il["Customer ID"] = clean_id_series(
        disbursement_sheet_il["Customer ID"]
    )

    disbursement_sheet_il["Account"] = clean_account_series(
        disbursement_sheet_il["Account"]
    )

    disbursement_sheet_il["IFSC Code"] = clean_ifsc_series(
        disbursement_sheet_il["IFSC Code"]
    )

    # ARC
    arc_sheet["Member ID"] = clean_id_series(
        arc_sheet["Member ID"]
    )

    # ========================================================
    # NUMERIC COLUMNS
    # ========================================================

    icici_sheet["Amount"] = pd.to_numeric(
        icici_sheet["Amount"]
        .str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

    icici_sheet_il["Amount"] = pd.to_numeric(
        icici_sheet_il["Amount"]
        .str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

    disbursement_sheet["Final Disbursal Amount"] = pd.to_numeric(
        disbursement_sheet["Final Disbursal Amount"]
        .str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

    disbursement_sheet_il["Final Disbursal Amount"] = pd.to_numeric(
        disbursement_sheet_il["Final Disbursal Amount"]
        .str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

    # ========================================================
    # DATE CLEANING
    # ========================================================

    master_sheet["Posting-Date"] = parse_date_column(
        master_sheet["Posting-Date"]
    )

    # ========================================================
    # COMBINED ICICI
    # ========================================================

    icici_combined = pd.concat(
        [
            icici_sheet,
            icici_sheet_il
        ],
        ignore_index=True
    )

    # ========================================================
    # DISB CALCULATION
    # ========================================================

    # Your old code called this "four_months_ago"
    # but actually used 90 days.
    # We are retaining the SAME 90-day rule.

    back_date = pd.Timestamp(
        datetime.now() - timedelta(days=90)
    ).normalize()

    # ========================================================
    # 1. ICICI VS MASTER - CUSTOMER ID MATCH
    # ========================================================

    merged_df1 = pd.merge(
        icici_combined,
        master_sheet,
        on="Coustmer ID",
        how="inner"
    )

    output1 = merged_df1[
        [
            "Coustmer ID",
            "A/C no._x",
            "A/C no._y",
            "Posting-Date",
        ]
    ].copy()

    output1["90_Days_Back_Date"] = back_date.date()

    output1["A/C_No_Condition"] = (
        output1["A/C no._x"]
        ==
        output1["A/C no._y"]
    )

    output1["Posting_Date_Condition"] = (
        output1["Posting-Date"]
        >=
        back_date
    )

    final_output1_1 = output1[
        (
            output1["A/C_No_Condition"]
        )
        &
        (
            output1["Posting_Date_Condition"]
        )
    ].copy()

    # ========================================================
    # 2. ICICI VS MASTER - ACCOUNT NUMBER MATCH
    # ========================================================

    merged_df1_1 = pd.merge(
        icici_combined,
        master_sheet,
        on="A/C no.",
        how="inner"
    )

    output1_1 = merged_df1_1[
        [
            "A/C no.",
            "Coustmer ID_x",
            "Coustmer ID_y",
            "Posting-Date",
        ]
    ].copy()

    output1_1["Posting_Date_Condition"] = (
        output1_1["Posting-Date"]
        >=
        back_date
    )

    final_output1_2 = output1_1[
        output1_1["Posting_Date_Condition"]
    ].copy()

    final_output1 = pd.concat(
        [
            final_output1_1,
            final_output1_2
        ],
        ignore_index=True
    )

    # ========================================================
    # 3. ICICI JLG VS DB JLG
    # ========================================================

    merged_df2 = pd.merge(
        icici_sheet,
        disbursement_sheet,
        left_on="Loan Account ID",
        right_on="Loan ID",
        how="inner"
    )

    output2 = merged_df2[
        [
            "Loan Account ID",
            "Loan ID",
            "A/C no.",
            "Account",
            "IFSC Code_x",
            "IFSC Code_y",
            "Coustmer ID",
            "Customer ID",
            "Final Disbursal Amount",
            "Amount",
        ]
    ].copy()

    output2["A/C_No_Condition"] = (
        output2["A/C no."]
        ==
        output2["Account"]
    )

    output2["IFSC_Condition"] = (
        output2["IFSC Code_x"]
        ==
        output2["IFSC Code_y"]
    )

    output2["Amount_Difference"] = (
        output2["Final Disbursal Amount"]
        -
        output2["Amount"]
    )

    final_output2 = output2[
        (
            output2["A/C_No_Condition"]
        )
        &
        (
            output2["IFSC_Condition"]
        )
    ].copy()

    final_output3 = output2[
        (
            ~output2["A/C_No_Condition"]
        )
        |
        (
            ~output2["IFSC_Condition"]
        )
    ].copy()

    # ========================================================
    # 4. ICICI IL VS DB IL
    # ========================================================

    merged_df4 = pd.merge(
        icici_sheet_il,
        disbursement_sheet_il,
        left_on="Loan Account ID",
        right_on="Loan ID",
        how="inner"
    )

    output4 = merged_df4[
        [
            "Loan Account ID",
            "Loan ID",
            "A/C no.",
            "Account",
            "IFSC Code_x",
            "IFSC Code_y",
            "Coustmer ID",
            "Customer ID",
            "Final Disbursal Amount",
            "Amount",
        ]
    ].copy()

    output4["A/C_No_Condition"] = (
        output4["A/C no."]
        ==
        output4["Account"]
    )

    output4["IFSC_Condition"] = (
        output4["IFSC Code_x"]
        ==
        output4["IFSC Code_y"]
    )

    output4["Amount_Difference"] = (
        output4["Final Disbursal Amount"]
        -
        output4["Amount"]
    )

    final_output4 = output4[
        (
            output4["A/C_No_Condition"]
        )
        &
        (
            output4["IFSC_Condition"]
        )
    ].copy()

    final_output5 = output4[
        (
            ~output4["A/C_No_Condition"]
        )
        |
        (
            ~output4["IFSC_Condition"]
        )
    ].copy()

    # ========================================================
    # 5. ICICI VS ARC
    # ========================================================

    merged_df6 = pd.merge(
        icici_combined,
        arc_sheet,
        left_on="Coustmer ID",
        right_on="Member ID",
        how="inner"
    )

    output6 = merged_df6[
        [
            "Coustmer ID"
        ]
    ].copy()

    # ========================================================
    # 6. DB JLG NOT PRESENT IN ICICI
    # ========================================================

    merged_df7 = disbursement_sheet.merge(
        icici_sheet,
        left_on="Loan ID",
        right_on="Loan Account ID",
        how="left",
        indicator=True,
        suffixes=("_DB", "_ICICI")
    )

    unique_db_jlg = merged_df7[
        merged_df7["_merge"] == "left_only"
    ].drop(
        columns=["_merge"]
    ).copy()

    # ========================================================
    # 7. DB IL NOT PRESENT IN ICICI
    # ========================================================

    merged_df8 = disbursement_sheet_il.merge(
        icici_sheet_il,
        left_on="Loan ID",
        right_on="Loan Account ID",
        how="left",
        indicator=True,
        suffixes=("_DB", "_ICICI")
    )

    unique_db_il = merged_df8[
        merged_df8["_merge"] == "left_only"
    ].drop(
        columns=["_merge"]
    ).copy()

    # ========================================================
    # PARIVAAR CALCULATION
    # ========================================================

    # -----------------------------
    # JLG
    # -----------------------------

    parivaar_jlg_merge = pd.merge(
        icici_sheet,
        disbursement_sheet,
        left_on="Loan Account ID",
        right_on="Loan ID",
        how="inner",
        suffixes=("_ICICI", "_DB")
    )

    product_column_jlg = get_product_column(
        parivaar_jlg_merge
    )

    # Because suffixes above may generate Product_ICICI/Product_DB
    if product_column_jlg is None:

        for possible in [
            "Product_ICICI",
            "Product_DB"
        ]:
            if possible in parivaar_jlg_merge.columns:
                product_column_jlg = possible
                break

    if product_column_jlg is None:
        raise ValueError(
            "Product column was not found for the JLG "
            "Parivaar calculation. Please check ICICI.xlsx "
            "and CASHLESS DISBURSEMENT.xlsx."
        )

    parivaar_jlg = parivaar_jlg_merge[
        parivaar_jlg_merge[
            product_column_jlg
        ]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .str.startswith("parivaar")
    ].copy()

    parivaar_jlg["Source"] = "JLG"

    # -----------------------------
    # IL
    # -----------------------------

    parivaar_il_merge = pd.merge(
        icici_sheet_il,
        disbursement_sheet_il,
        left_on="Loan Account ID",
        right_on="Loan ID",
        how="inner",
        suffixes=("_ICICI", "_DB")
    )

    product_column_il = get_product_column(
        parivaar_il_merge
    )

    if product_column_il is None:

        for possible in [
            "Product_ICICI",
            "Product_DB"
        ]:
            if possible in parivaar_il_merge.columns:
                product_column_il = possible
                break

    if product_column_il is None:
        raise ValueError(
            "Product column was not found for the IL "
            "Parivaar calculation. Please check ICICI IL.xlsx "
            "and CASHLESS DISBURSEMENT IL.xlsx."
        )

    parivaar_il = parivaar_il_merge[
        parivaar_il_merge[
            product_column_il
        ]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .str.startswith("parivaar")
    ].copy()

    parivaar_il["Source"] = "IL"

    all_parivaar = pd.concat(
        [
            parivaar_jlg,
            parivaar_il
        ],
        ignore_index=True
    )

    # ========================================================
    # DB VS ARC CALCULATION
    # ========================================================

    arc_member_ids = set(
        arc_sheet["Member ID"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    arc_jlg = disbursement_sheet[
        disbursement_sheet[
            "Customer ID"
        ].isin(
            arc_member_ids
        )
    ].copy()

    arc_jlg["Source"] = "DB JLG"

    arc_il = disbursement_sheet_il[
        disbursement_sheet_il[
            "Customer ID"
        ].isin(
            arc_member_ids
        )
    ].copy()

    arc_il["Source"] = "DB IL"

    all_arc = pd.concat(
        [
            arc_jlg,
            arc_il
        ],
        ignore_index=True
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    summary = pd.DataFrame(
        {
            "Particular": [
                "Records in ICICI JLG",
                "Records in ICICI IL",
                "Records in DB JLG",
                "Records in DB IL",

                "JLG Account No Matched",
                "IL Account No Matched",

                "JLG Fully Matched",
                "IL Fully Matched",

                "JLG False Cases",
                "IL False Cases",

                "ICICI vs ARC Cases",

                "DB JLG Not in ICICI",
                "DB IL Not in ICICI",

                "JLG Parivaar Cases",
                "IL Parivaar Cases",
                "Total Parivaar Cases",

                "JLG ARC Matches",
                "IL ARC Matches",
                "Total ARC Matches",
            ],

            "Count": [
                len(icici_sheet),
                len(icici_sheet_il),
                len(disbursement_sheet),
                len(disbursement_sheet_il),

                int(
                    output2[
                        "A/C_No_Condition"
                    ].sum()
                ),

                int(
                    output4[
                        "A/C_No_Condition"
                    ].sum()
                ),

                len(final_output2),
                len(final_output4),

                len(final_output3),
                len(final_output5),

                len(output6),

                len(unique_db_jlg),
                len(unique_db_il),

                len(parivaar_jlg),
                len(parivaar_il),
                len(all_parivaar),

                len(arc_jlg),
                len(arc_il),
                len(all_arc),
            ],
        }
    )

    # ========================================================
    # RETURN OUTPUT
    # ========================================================

    return {
        "Summary": summary,

        "ICICI_vs_Master": final_output1,

        "JLG_Matched": final_output2,
        "JLG_False": final_output3,

        "IL_Matched": final_output4,
        "IL_False": final_output5,

        "ICICI_vs_ARC": output6,

        "JLG_DB_Not_ICICI": unique_db_jlg,
        "IL_DB_Not_ICICI": unique_db_il,

        "Parivaar_JLG": parivaar_jlg,
        "Parivaar_IL": parivaar_il,
        "Parivaar_All": all_parivaar,

        "ARC_JLG": arc_jlg,
        "ARC_IL": arc_il,
        "ARC_All": all_arc,
    }
