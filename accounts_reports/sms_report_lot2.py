import os
import zipfile
import pandas as pd


def process_sms_report_lot2(files, output_dir, progress_callback=None):
    os.makedirs(output_dir, exist_ok=True)

    if progress_callback:
        progress_callback(10, "Reading uploaded files...")

    output_files = []

    for report_name, uploaded_file in files.items():
        uploaded_file.seek(0)

        df = pd.read_excel(uploaded_file)
        out_path = os.path.join(output_dir, f"{report_name}.csv")

        df.to_csv(out_path, index=False, encoding="utf-8-sig")
        output_files.append(out_path)

    if progress_callback:
        progress_callback(90, "Creating ZIP file...")

    zip_path = os.path.join(output_dir, "SMS_Report_Lot2_Output.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file_path in output_files:
            zipf.write(file_path, arcname=os.path.basename(file_path))

    if progress_callback:
        progress_callback(100, "Completed.")

    summary = [
        {
            "report_name": name,
            "rows": len(pd.read_csv(path))
        }
        for name, path in zip(files.keys(), output_files)
    ]

    return summary, zip_path
