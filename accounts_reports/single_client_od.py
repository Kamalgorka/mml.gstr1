import os
import tempfile
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side


EXCLUDE_FUNDERS = {
    "ARCILARC_MARCH_2026",
    "CFMARC_MARCH_2025",
    "PHOENIX_ARC",
    "PHOENIX_ARC-1",
}


HUB_MAPPING = {
    "Lucknow_Hub": "Hub-4 Lucknow",
    "Patna_Hub": "Hub-3 Patna",
    "Chandigarh_Hub": "Hub-1 Chandigarh",
    "Bhubaneswar_Hub": "Hub-6 Bhubaneswar",
    "Ahmedabad_Hub": "Hub-2 Ahmedabad",
    "Kolkata_Hub": "Hub-5 Kolkata",
}


def read_uploaded_file(uploaded_file):
    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    if filename.endswith(".xlsb"):
        return pd.read_excel(uploaded_file, engine="pyxlsb")

    return pd.read_excel(uploaded_file)


def process_single_client_od(uploaded_file):

    df = read_uploaded_file(uploaded_file)

    # ------------------------------------
    # Data Cleaning
    # ------------------------------------

    df["status"] = df["status"].astype(str).str.strip()
    df["funder"] = df["funder"].astype(str).str.strip()

    df["od_days"] = pd.to_numeric(
        df["od_days"],
        errors="coerce"
    )

    # Remove Death
    df = df[
        ~df["status"].str.upper().eq("DEATH")
    ]

    # Remove Funders
    df = df[
        ~df["funder"].str.upper().isin(
            {x.upper() for x in EXCLUDE_FUNDERS}
        )
    ]

    # OD Days <= 90
    df = df[
        df["od_days"] <= 90
    ]

    # ------------------------------------
    # Single Client Centers
    # ------------------------------------

    center_counts = (
        df.groupby("center_name")["cust_id"]
        .nunique()
    )

    single_centers = center_counts[
        center_counts == 1
    ].index

    final_df = df[
        df["center_name"].isin(single_centers)
    ].copy()

    # ------------------------------------
    # Hub Mapping
    # ------------------------------------

    final_df["Hub_Name"] = (
        final_df["zone_name"]
        .map(HUB_MAPPING)
        .fillna(final_df["zone_name"])
    )

    # ------------------------------------
    # Summary Data
    # ------------------------------------

    state_summary = (
        final_df.groupby(
            ["Hub_Name", "state_name"],
            dropna=False
        )
        .agg(
            Arrear_Amount=("total_arrear", "sum"),
            PAR_Amount=("outstanding_principal", "sum"),
            No_of_Centers=("center_name", pd.Series.nunique),
            No_of_Clients=("cust_id", pd.Series.nunique),
        )
        .reset_index()
    )

    hub_totals = (
        final_df.groupby("Hub_Name")
        .agg(
            Arrear_Amount=("total_arrear", "sum"),
            PAR_Amount=("outstanding_principal", "sum"),
            No_of_Centers=("center_name", pd.Series.nunique),
            No_of_Clients=("cust_id", pd.Series.nunique),
        )
        .reset_index()
    )

    hub_order = [
        "Hub-1 Chandigarh",
        "Hub-2 Ahmedabad",
        "Hub-3 Patna",
        "Hub-4 Lucknow",
        "Hub-5 Kolkata",
        "Hub-6 Bhubaneswar",
    ]

    summary_rows = []

    summary_rows.append([
        "Single Client OD in a Center (Non-NPA Clients)",
        "",
        "",
        "",
        "",
    ])

    summary_rows.append([
        "HUB/ Region",
        "Arrear Amount",
        "PAR Amount",
        "No. of Centers",
        "No. of Clients",
    ])

    for hub in hub_order:

        hub_data = state_summary[
            state_summary["Hub_Name"] == hub
        ]

        if hub_data.empty:
            continue

        summary_rows.append([hub, "", "", "", ""])

        for _, row in hub_data.iterrows():

            summary_rows.append([
                row["state_name"],
                row["Arrear_Amount"],
                row["PAR_Amount"],
                row["No_of_Centers"],
                row["No_of_Clients"],
            ])

        total_row = hub_totals[
            hub_totals["Hub_Name"] == hub
        ].iloc[0]

        summary_rows.append([
            f"{hub} Total",
            total_row["Arrear_Amount"],
            total_row["PAR_Amount"],
            total_row["No_of_Centers"],
            total_row["No_of_Clients"],
        ])

    summary_rows.append([
        "Grand Total",
        final_df["total_arrear"].sum(),
        final_df["outstanding_principal"].sum(),
        final_df["center_name"].nunique(),
        final_df["cust_id"].nunique(),
    ])

    summary_df = pd.DataFrame(summary_rows)

    # ------------------------------------
    # Save Excel
    # ------------------------------------

    output_dir = tempfile.mkdtemp()

    output_path = os.path.join(
        output_dir,
        "Single Client OD in a Center.xlsx"
    )

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
            header=False
        )

        final_df.to_excel(
            writer,
            sheet_name="Member wise data",
            index=False
        )

    # ------------------------------------
    # Formatting
    # ------------------------------------

    wb = load_workbook(output_path)

    ws = wb["Summary"]

    dark_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    title_fill = PatternFill(
        "solid",
        fgColor="B7DEE8"
    )

    hub_fill = PatternFill(
        "solid",
        fgColor="D9EAD3"
    )

    state_fill = PatternFill(
        "solid",
        fgColor="E2F0D9"
    )

    total_fill = PatternFill(
        "solid",
        fgColor="D9EAD3"
    )

    grand_fill = PatternFill(
        "solid",
        fgColor="B7DEE8"
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = dark_border

    for cell in ws[1]:
        cell.fill = title_fill
        cell.font = Font(bold=True)

    for cell in ws[2]:
        cell.fill = title_fill
        cell.font = Font(bold=True)

    for row in range(3, ws.max_row + 1):

        value = str(ws.cell(row, 1).value)

        if value.startswith("Hub-") and "Total" not in value:
            for cell in ws[row]:
                cell.fill = hub_fill
                cell.font = Font(bold=True)

        elif value.startswith("Hub-") and "Total" in value:
            for cell in ws[row]:
                cell.fill = total_fill
                cell.font = Font(bold=True)

        elif value == "Grand Total":
            for cell in ws[row]:
                cell.fill = grand_fill
                cell.font = Font(bold=True)

        else:
            for cell in ws[row]:
                cell.fill = state_fill

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15

    wb.save(output_path)

    return output_path
