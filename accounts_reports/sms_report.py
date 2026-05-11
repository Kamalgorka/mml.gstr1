else:

    import tempfile
    import os
    from accounts_reports.sms_report import process_sms_report

    with st.spinner("Processing SMS Reports..."):

        with tempfile.TemporaryDirectory() as workdir:

            summary, zip_path = process_sms_report(
                files=required_files,
                output_dir=workdir
            )

            with open(zip_path, "rb") as f:
                zip_bytes = f.read()

    st.success("SMS Report processed successfully. 16 output files created.")

    st.write("### Processing Summary")

    for item in summary:
        st.write(
            f"✅ {item['report_name']} | "
            f"Arrear Free: {item['arrear_free_rows']} rows | "
            f"Arrear: {item['arrear_rows']} rows"
        )

    st.download_button(
        "⬇️ Download SMS Report Output ZIP",
        data=zip_bytes,
        file_name="SMS_Report_Output.zip",
        mime="application/zip",
        use_container_width=True
    )
