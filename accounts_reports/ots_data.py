import os
import tempfile
from datetime import datetime

import pandas as pd


def process_ots_data(uploaded_file, progress_callback=None):

    if progress_callback:
        progress_callback(10, "Reading OTS Data file...")

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)

    elif file_name.endswith(".xlsb"):
        df = pd.read_excel(uploaded_file, engine="pyxlsb")

    elif file_name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded_file)

    else:
        raise Exception("Unsupported file format.")

    if progress_callback:
        progress_callback(50, "Processing OTS Data...")

    df.columns = [str(c).strip() for c in df.columns]

    df["Processed_Date"] = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    if progress_callback:
        progress_callback(80, "Creating output file...")

    output_dir = tempfile.mkdtemp()
    output_file = os.path.join(output_dir, "OTS_Data_Output.xlsx")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="OTS Data", index=False)

    if progress_callback:
        progress_callback(100, "OTS Data report generated successfully.")

    return output_file
