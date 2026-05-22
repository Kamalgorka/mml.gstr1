import streamlit as st
from auth_utils import require_role, log_activity
from ui import load_global_css

# =========================================================
# PAGE CONFIG - must be first Streamlit command
# =========================================================
st.set_page_config(
    page_title="Accounts Reporting Team",
    page_icon="📘",
    layout="wide"
)

load_global_css()

# Hide default Streamlit sidebar navigation
st.markdown("""
<style>
[data-testid="stSidebarNav"] {
    display: none;
}
</style>
""", unsafe_allow_html=True)


# Sidebar Menu
def role_sidebar():
    role = st.session_state.get("user_role", "").strip().upper()

    st.sidebar.title("📌 Menu")
    st.sidebar.page_link("streamlit_app.py", label="Home", icon="🏠")

    if role in ["REPORTING", "ADMIN"]:
        st.sidebar.page_link(
            "pages/2_Reporting_Team.py",
            label="Reporting Team",
            icon="📊"
        )

    if role in ["HO", "ADMIN"]:
        st.sidebar.page_link(
            "pages/1_HO_Team.py",
            label="HO Team",
            icon="🏢"
        )

    if role in ["ACCOUNTS", "ADMIN"]:
        st.sidebar.page_link(
            "pages/3_Accounts_Reporting_Team.py",
            label="Accounts Reporting Team",
            icon="📘"
        )

    st.sidebar.markdown("---")

    if st.sidebar.button("🚪 Logout"):
        st.session_state.clear()
        st.switch_page("streamlit_app.py")


# Access Control
require_role(["ACCOUNTS", "ADMIN"])

# Load Sidebar
role_sidebar()

# Activity Log
log_activity("PAGE_OPEN", "Accounts Reporting Team")


# =========================================================
# MAIN PAGE
# =========================================================
st.title("📘 Accounts Reporting Team - Reports Automisation")

st.markdown(
    "Select report from below and upload required files to generate output."
)

report_options = [
    "Select Report",
    "1) SMS Report",
    "2) SMS Report Lot2",
    "3) ARC and Write Off Entries",
    "4) WriteOff Loan Collection",
    "5) OTS Data"
]

selected_report = st.selectbox(
    "🔎 Select Report (type to search)",
    report_options
)

st.markdown("---")


if selected_report == "Select Report":

    st.info("Please select a report to continue.")


elif selected_report == "1) SMS Report":

    st.subheader("📩 SMS Report Automation")

    st.caption(
        "Upload all required raw MIS files and run automation."
    )

    col1, col2 = st.columns(2)

    with col1:

        hub1_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 1",
            type=["xlsx", "xls"],
            key="hub1"
        )

        hub2_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 2",
            type=["xlsx", "xls"],
            key="hub2"
        )

        hub3_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 3",
            type=["xlsx", "xls"],
            key="hub3"
        )

        hub4_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 4",
            type=["xlsx", "xls"],
            key="hub4"
        )

    with col2:

        hub5_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 5",
            type=["xlsx", "xls"],
            key="hub5"
        )

        hub6_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 6",
            type=["xlsx", "xls"],
            key="hub6"
        )

        il_file = st.file_uploader(
            "Monthly Outstanding SMS Data IL",
            type=["xlsx", "xls"],
            key="il"
        )

        writeoff_file = st.file_uploader(
            "Loan OS Write Off",
            type=["xlsx", "xls"],
            key="writeoff"
        )

        writeoff_il_file = st.file_uploader(
            "Loan OS Write Off IL",
            type=["xlsx", "xls"],
            key="writeoff_il"
        )

    st.markdown("---")

    run_btn = st.button(
        "🚀 Run SMS Automation",
        use_container_width=True,
        key="run_sms_lot1"
    )

    if run_btn:

        st.warning("Run button clicked. Validation started...")

        required_files = {
            "Monthly Outstanding SMS Data JLG HUB 1": hub1_file,
            "Monthly Outstanding SMS Data JLG HUB 2": hub2_file,
            "Monthly Outstanding SMS Data JLG HUB 3": hub3_file,
            "Monthly Outstanding SMS Data JLG HUB 4": hub4_file,
            "Monthly Outstanding SMS Data JLG HUB 5": hub5_file,
            "Monthly Outstanding SMS Data JLG HUB 6": hub6_file,
            "Monthly Outstanding SMS Data IL": il_file,
            "Loan OS Write Off": writeoff_file,
            "Loan OS Write Off IL": writeoff_il_file,
        }

        missing = [
            name for name, file in required_files.items()
            if file is None
        ]

        if missing:

            st.error("Please upload all required files.")

            for m in missing:
                st.write(f"❌ {m}")

        else:

            try:
                import tempfile
                from accounts_reports.sms_report import process_sms_report

                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_progress(percent, message):
                    progress_bar.progress(percent)
                    status_box.write(f"⏳ {percent}% - {message}")

                with st.spinner("Processing SMS Reports..."):

                    with tempfile.TemporaryDirectory() as workdir:

                        summary, zip_path = process_sms_report(
                            files=required_files,
                            output_dir=workdir,
                            progress_callback=update_progress
                        )

                        with open(zip_path, "rb") as f:
                            zip_bytes = f.read()

                progress_bar.progress(100)
                status_box.success("✅ 100% - SMS Report processing completed.")

                st.success("SMS Report processed successfully.")

                st.write("### Processing Summary")

                for item in summary:

                    if item.get("type") == "monthly":
                        st.write(
                            f"✅ {item['report_name']} | "
                            f"Arrear Free: {item['arrear_free_rows']} rows | "
                            f"Arrear: {item['arrear_rows']} rows | "
                            f"Arrear Free Discrepancies: {item['arrear_free_discrepancy_rows']} | "
                            f"Arrear Discrepancies: {item['arrear_discrepancy_rows']}"
                        )

                    elif item.get("type") == "writeoff":
                        st.write(
                            f"✅ {item['report_name']} | "
                            f"WriteOff rows: {item['total_writeoff_rows_after_filter']} | "
                            f"Unique Cust IDs: {item['unique_cust_id_rows']} | "
                            f"Discrepancies: {item['discrepancy_rows']}"
                        )

                st.download_button(
                    "⬇️ Download SMS Report Output ZIP",
                    data=zip_bytes,
                    file_name="SMS_Report_Output.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Error occurred while processing SMS Report.")
                st.exception(e)


elif selected_report == "2) SMS Report Lot2":

    st.subheader("📩 SMS Report Lot2 Automation")

    st.caption(
        "Upload all required raw files and run Lot2 automation."
    )

    col1, col2 = st.columns(2)

    with col1:

        not_sent_sms_file = st.file_uploader(
            "Not Sent SMS Data",
            type=["xlsx", "xls"],
            key="lot2_not_sent_sms"
        )

        not_sent_writeoff_sms_file = st.file_uploader(
            "Not Sent Write Off SMS Data",
            type=["xlsx", "xls"],
            key="lot2_not_sent_writeoff_sms"
        )

    with col2:

        loan_os_writeoff_file = st.file_uploader(
            "Loan OS Write Off",
            type=["xlsx", "xls"],
            key="lot2_loan_os_writeoff"
        )

        loan_os_writeoff_il_file = st.file_uploader(
            "Loan OS Write Off IL",
            type=["xlsx", "xls"],
            key="lot2_loan_os_writeoff_il"
        )

    st.markdown("---")

    run_btn = st.button(
        "🚀 Run SMS Lot2 Automation",
        use_container_width=True,
        key="run_sms_lot2"
    )

    if run_btn:

        required_files = {
            "Not Sent SMS Data": not_sent_sms_file,
            "Not Sent Write Off SMS Data": not_sent_writeoff_sms_file,
            "Loan OS Write Off": loan_os_writeoff_file,
            "Loan OS Write Off IL": loan_os_writeoff_il_file,
        }

        missing = [
            name for name, file in required_files.items()
            if file is None
        ]

        if missing:

            st.error("Please upload all required files.")

            for m in missing:
                st.write(f"❌ {m}")

        else:

            try:
                import tempfile
                from accounts_reports.sms_report_lot2 import process_sms_report_lot2

                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_progress(percent, message):
                    progress_bar.progress(percent)
                    status_box.write(f"⏳ {percent}% - {message}")

                with st.spinner("Processing SMS Report Lot2..."):

                    with tempfile.TemporaryDirectory() as workdir:

                        summary, zip_path = process_sms_report_lot2(
                            files=required_files,
                            output_dir=workdir,
                            progress_callback=update_progress
                        )

                        with open(zip_path, "rb") as f:
                            zip_bytes = f.read()

                progress_bar.progress(100)
                status_box.success("✅ 100% - SMS Report Lot2 processing completed.")

                st.success("SMS Report Lot2 processed successfully.")

                st.write("### Processing Summary")

                for item in summary:
                    if "discrepancy_rows" in item:
                        st.write(
                            f"✅ {item['report_name']} | "
                            f"Rows: {item['rows']} | "
                            f"Discrepancies: {item['discrepancy_rows']}"
                        )
                    else:
                        st.write(
                            f"✅ {item['report_name']} | "
                            f"Rows: {item['rows']}"
                        )

                st.download_button(
                    "⬇️ Download SMS Report Lot2 Output ZIP",
                    data=zip_bytes,
                    file_name="SMS_Report_Lot2_Output.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Error occurred while processing SMS Report Lot2.")
                st.exception(e)


elif selected_report == "3) ARC and Write Off Entries":

    st.subheader("📘 ARC and Write Off Entries")

    st.caption(
        "Upload ARC / Write Off Excel file. System will prepare separate entry file for each sheet and one consolidated file."
    )

    arc_file = st.file_uploader(
        "Upload ARC / Write Off Entries File",
        type=["xlsx", "xls"],
        key="arc_writeoff_file"
    )

    st.markdown("---")

    run_btn = st.button(
        "🚀 Run ARC and Write Off Entries",
        use_container_width=True,
        key="run_arc_writeoff"
    )

    if run_btn:

        if arc_file is None:
            st.error("Please upload ARC / Write Off file.")

        else:

            try:
                import tempfile
                from accounts_reports.arc_writeoff_entries import process_arc_writeoff_entries

                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_progress(percent, message):
                    progress_bar.progress(percent)
                    status_box.write(f"⏳ {percent}% - {message}")

                with st.spinner("Processing ARC and Write Off Entries..."):

                    with tempfile.TemporaryDirectory() as workdir:

                        summary, zip_path = process_arc_writeoff_entries(
                            uploaded_file=arc_file,
                            output_dir=workdir,
                            progress_callback=update_progress
                        )

                        with open(zip_path, "rb") as f:
                            zip_bytes = f.read()

                progress_bar.progress(100)
                status_box.success("✅ 100% - ARC and Write Off Entries completed.")

                st.success("ARC and Write Off Entries processed successfully.")

                st.write("### Processing Summary")

                for item in summary:
                    st.write(
                        f"✅ {item['sheet_name']} | "
                        f"Rows: {item['rows']}"
                    )

                st.download_button(
                    "⬇️ Download ARC and Write Off Entries ZIP",
                    data=zip_bytes,
                    file_name="ARC_WriteOff_Entries_Output.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Error occurred while processing ARC and Write Off Entries.")
                st.exception(e)


elif selected_report == "4) WriteOff Loan Collection":

    st.subheader("📘 WriteOff Loan Collection")

    st.caption(
        "Upload WriteOff Loan Collection file and Repayment file. "
        "System will update System Write-Off and Manual Write-Off sheets month-wise."
    )

    col1, col2 = st.columns(2)

    with col1:
        writeoff_collection_file = st.file_uploader(
            "WriteOff Loan Collection File",
            type=["xlsx", "xls"],
            key="writeoff_collection_file"
        )

    with col2:
        repayment_file = st.file_uploader(
            "Repayment File",
            type=["xlsx", "xls", "xlsb"],
            key="writeoff_collection_repayment_file"
        )

    st.markdown("---")

    run_btn = st.button(
        "🚀 Run WriteOff Loan Collection",
        use_container_width=True,
        key="run_writeoff_loan_collection"
    )

    if run_btn:

        required_files = {
            "WriteOff Loan Collection File": writeoff_collection_file,
            "Repayment File": repayment_file,
        }

        missing = [
            name for name, file in required_files.items()
            if file is None
        ]

        if missing:

            st.error("Please upload all required files.")

            for m in missing:
                st.write(f"❌ {m}")

        else:

            try:
                import os
                import tempfile
                from accounts_reports.writeoff_loan_collection import process_writeoff_loan_collection

                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_progress(percent, message):
                    progress_bar.progress(percent)
                    status_box.write(f"⏳ {percent}% - {message}")

                with st.spinner("Processing WriteOff Loan Collection..."):

                    with tempfile.TemporaryDirectory() as workdir:

                        output_file = process_writeoff_loan_collection(
                            writeoff_file=writeoff_collection_file,
                            repayment_file=repayment_file,
                            output_dir=workdir,
                            progress_callback=update_progress
                        )

                        with open(output_file, "rb") as f:
                            output_bytes = f.read()

                progress_bar.progress(100)
                status_box.success("✅ 100% - WriteOff Loan Collection completed.")

                st.success("WriteOff Loan Collection processed successfully.")

                st.download_button(
                    "⬇️ Download Updated WriteOff Loan Collection",
                    data=output_bytes,
                    file_name=os.path.basename(output_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Error occurred while processing WriteOff Loan Collection.")
                st.exception(e)

elif selected_report == "5) OTS Data":

    st.subheader("📘 OTS Data")

    st.caption(
        "Upload OTS Data file. System will process the file and generate output."
    )

    ots_file = st.file_uploader(
        "Upload OTS Data File",
        type=["xlsx", "xls", "xlsb", "csv"],
        key="ots_data_file"
    )

    st.markdown("---")

    run_btn = st.button(
        "🚀 Run OTS Data",
        use_container_width=True,
        key="run_ots_data"
    )

    if run_btn:

        if ots_file is None:
            st.error("Please upload OTS Data file.")

        else:

            try:
                from accounts_reports.ots_data import process_ots_data

                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_progress(percent, message):
                    progress_bar.progress(percent)
                    status_box.write(f"⏳ {percent}% - {message}")

                with st.spinner("Processing OTS Data..."):

                    output_file = process_ots_data(
                        uploaded_file=ots_file,
                        progress_callback=update_progress
                    )

                    with open(output_file, "rb") as f:
                        output_bytes = f.read()

                progress_bar.progress(100)
                status_box.success("✅ 100% - OTS Data completed.")

                st.success("OTS Data processed successfully.")

                st.download_button(
                    "⬇️ Download OTS Data Output",
                    data=output_bytes,
                    file_name="OTS_Data_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error("Error occurred while processing OTS Data.")
                st.exception(e)
