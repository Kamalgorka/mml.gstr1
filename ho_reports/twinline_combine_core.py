from io import BytesIO

import pandas as pd


OUTPUT_COLUMNS = [
    "Transaction DateTime",
    "Created On",
    "Amount",
    "Branch",
    "Loan ID",
    "Account ID Match",
    "Payer VPA",
    "Payer Mobile No.",
    "Cust Ref",
    "Upi TxnId",
    "Recon Msg",
]


JLG_COLUMNS = {
    "Transaction DateTime": "Transaction DateTime",
    "Created On": "Created On",
    "Amount": "Amount",
    "Branch": "Branch",
    "Loan ID": "Loan ID",
    "Account ID Match": "Account ID Match",
    "Payer VPA": "Payer VPA",
    "Payer Mobile No.": "Payer Mobile No.",
    "Cust Ref": "Cust Ref",
    "Upi TxnId": "Upi TxnId",
    "Recon Msg": "Recon Msg",
}


IL_COLUMNS = {
    "Transaction DateTime": "Transaction DateTime",
    "Created On": "Created On",
    "Amount": "Amount",
    "Branch": "Branch",
    "Account ID": "Loan ID",
    "Account Id Match": "Account ID Match",
    "Payer VPA": "Payer VPA",
    "Payer Mobile No.": "Payer Mobile No.",
    "Cust Ref": "Cust Ref",
    "Upi TxnId": "Upi TxnId",
    "Recon Msg": "Recon Msg",
}


def clean_headers(df):
    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


def validate_required_columns(
    df,
    required_columns,
    file_name
):
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


def prepare_dataframe(
    df,
    mapping,
    file_name
):
    df = clean_headers(df)

    validate_required_columns(
        df,
        list(mapping.keys()),
        file_name
    )

    selected = (
        df[list(mapping.keys())]
        .rename(columns=mapping)
        .copy()
    )

    return selected


def generate_twinline_report(
    jlg_file,
    il_file
):

    # ========================================================
    # READ BOTH FILES
    # ========================================================

    df_jlg = pd.read_excel(
        jlg_file,
        engine="openpyxl"
    )

    df_il = pd.read_excel(
        il_file,
        engine="openpyxl"
    )

    # ========================================================
    # PREPARE JLG
    # ========================================================

    df_jlg_selected = prepare_dataframe(
        df_jlg,
        JLG_COLUMNS,
        "JLG.xlsx"
    )

    # ========================================================
    # PREPARE IL
    # ========================================================

    df_il_selected = prepare_dataframe(
        df_il,
        IL_COLUMNS,
        "IL.xlsx"
    )

    # ========================================================
    # SAME COLUMN ORDER
    # ========================================================

    df_jlg_selected = df_jlg_selected[
        OUTPUT_COLUMNS
    ]

    df_il_selected = df_il_selected[
        OUTPUT_COLUMNS
    ]

    # ========================================================
    # COMBINE BOTH
    # ========================================================

    combined_df = pd.concat(
        [
            df_jlg_selected,
            df_il_selected
        ],
        ignore_index=True
    )

    # ========================================================
    # BACKUP COLUMNS
    # ========================================================

    combined_df["Cust Ref 2"] = (
        combined_df["Cust Ref"]
        .fillna("")
        .astype(str)
    )

    combined_df["Payer Mobile No. 2"] = (
        combined_df["Payer Mobile No."]
        .fillna("")
        .astype(str)
    )

    # ========================================================
    # REMOVE DUPLICATES
    # Same Cust Ref + Amount
    # ========================================================

    combined_no_duplicates = (
        combined_df
        .drop_duplicates(
            subset=[
                "Cust Ref",
                "Amount"
            ],
            keep="first"
        )
        .reset_index(drop=True)
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    summary = {
        "jlg_records":
            len(df_jlg_selected),

        "il_records":
            len(df_il_selected),

        "combined_records":
            len(combined_df),

        "duplicate_records_removed":
            len(combined_df)
            - len(combined_no_duplicates),

        "final_records":
            len(combined_no_duplicates),
    }

    # ========================================================
    # CREATE EXCEL IN MEMORY
    # ========================================================

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        combined_no_duplicates.to_excel(
            writer,
            sheet_name="Twinline Combined",
            index=False
        )

        worksheet = writer.book[
            "Twinline Combined"
        ]

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        # ====================================================
        # AUTO WIDTH
        # ====================================================

        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = (
                column_cells[0]
                .column_letter
            )

            for cell in column_cells:

                value = (
                    ""
                    if cell.value is None
                    else str(cell.value)
                )

                max_length = max(
                    max_length,
                    len(value)
                )

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(
                    max_length + 2,
                    12
                ),
                40
            )

    output.seek(0)

    return (
        output.getvalue(),
        summary
    )
