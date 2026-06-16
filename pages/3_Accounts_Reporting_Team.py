import streamlit as st
from auth_utils import require_role, log_activity
from ui import load_global_css
from accounts_reports.ots_data import process_ots_data
from accounts_reports.single_client_od import process_single_client_od
from accounts_reports.od_amount_less_1000 import process_od_amount_less_1000
from accounts_reports.od_status_two_loans import process_od_status_two_loans
from accounts_reports.midfin_finpage_recon import process_midfin_finpage_recon
# =========================================================
# PAGE CONFIG - must be first Streamlit command
# =========================================================
st.set_page_config(
    page_title="Accounts Reporting Team",
    page_icon="📘",
    layout="wide"
)

load_global_css()

st.markdown("""
<style>
/* Selectbox main closed field */
div[data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1.5px solid #ff6b35 !important;
    border-radius: 8px !important;
    box-shadow: none !important;
}

/* Selected value text */
div[data-baseweb="select"] div,
div[data-baseweb="select"] span {
    color: #111827 !important;
    font-weight: 700 !important;
    opacity: 1 !important;
}

/* Dropdown menu container */
div[data-baseweb="popover"] {
    opacity: 1 !important;
}

/* Dropdown options */
div[role="option"] {
    color: #111827 !important;
    font-weight: 700 !important;
    opacity: 1 !important;
    background-color: #ffffff !important;
}

/* Text inside dropdown options */
div[role="option"] div,
div[role="option"] span {
    color: #111827 !important;
    font-weight: 700 !important;
    opacity: 1 !important;
}

/* Hover / selected row */
div[role="option"]:hover {
    background-color: #fff3ed !important;
}
</style>
""", unsafe_allow_html=True)
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
    "5) OTS Data",
    "6) Single Client OD in a Center",
    "7) OD Amount Less 1000 Non NPA Clients",
    "8) OD Status of Two Loans",
    "9) Midfin and Finpage Collection Recon"
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

        st.session_state.pop("sms_lot2_zip_bytes", None)
        st.session_state.pop("sms_lot2_zip_name", None)
        st.session_state.pop("sms_lot2_download_url", None)
        st.session_state.pop("sms_lot2_summary", None)

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
                import os
                import shutil
                from accounts_reports.sms_report_lot2 import process_sms_report_lot2

                progress_bar = st.progress(0)
                status_box = st.empty()

                def update_progress(percent, message):
                    progress_bar.progress(percent)
                    status_box.write(f"⏳ {percent}% - {message}")

                with st.spinner("Processing SMS Report Lot2..."):

                    workdir = "/tmp/mml_sms_lot2_outputs"

                    if os.path.exists(workdir):
                        shutil.rmtree(workdir)

                    os.makedirs(workdir, exist_ok=True)

                    summary, zip_path = process_sms_report_lot2(
                        files=required_files,
                        output_dir=workdir,
                        progress_callback=update_progress
                    )

                    with open(zip_path, "rb") as f:
                        st.session_state["sms_lot2_zip_bytes"] = f.read()
                    import uuid

                    download_dir = "/opt/apps/streamlit/mml_smart_reports/downloads/sms_lot2"
                    os.makedirs(download_dir, exist_ok=True)
                    
                    download_filename = f"SMS_Report_Lot2_Output_{uuid.uuid4().hex}.xlsx"
                    download_path = os.path.join(download_dir, download_filename)
                    
                    shutil.copy2(zip_path, download_path)
                    
                    st.session_state["sms_lot2_download_url"] = (
                        f"/mmlsmartreports/downloads/sms_lot2/{download_filename}"
                    )

                    st.session_state["sms_lot2_zip_name"] = "SMS_Report_Lot2_Output.zip"
                    st.session_state["sms_lot2_summary"] = summary

                progress_bar.progress(100)
                status_box.success("✅ 100% - SMS Report Lot2 processing completed.")

                st.success("SMS Report Lot2 processed successfully.")

            except Exception as e:
                st.error("Error occurred while processing SMS Report Lot2.")
                st.exception(e)

    if "sms_lot2_summary" in st.session_state:
        st.write("### Processing Summary")

        for item in st.session_state["sms_lot2_summary"]:
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

        if "sms_lot2_download_url" in st.session_state:
            st.link_button(
                "⬇️ Download SMS Report Lot2 Output Excel",
                st.session_state["sms_lot2_download_url"],
                use_container_width=True
            )
        
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
    st.subheader("📘 OTS Data Report")

    st.info("Upload all four required files to generate the OTS Data output.")

    col1, col2 = st.columns(2)

    with col1:
        outstanding_il_file = st.file_uploader(
            "Upload Outstanding IL",
            type=["xlsx", "xls", "xlsb", "csv"],
            key="ots_outstanding_il"
        )

        writeoff_il_file = st.file_uploader(
            "Upload Write Off IL",
            type=["xlsx", "xls", "xlsb", "csv"],
            key="ots_writeoff_il"
        )

    with col2:
        outstanding_jlg_file = st.file_uploader(
            "Upload Outstanding JLG",
            type=["xlsx", "xls", "xlsb", "csv"],
            key="ots_outstanding_jlg"
        )

        writeoff_jlg_file = st.file_uploader(
            "Upload Write Off JLG",
            type=["xlsx", "xls", "xlsb", "csv"],
            key="ots_writeoff_jlg"
        )

    if st.button("Generate OTS Data Report", type="primary"):
        if not all([
            outstanding_il_file,
            outstanding_jlg_file,
            writeoff_il_file,
            writeoff_jlg_file,
        ]):
            st.error("Please upload all four files before generating the report.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(value, message):
                progress_bar.progress(value)
                status_text.info(message)

            try:
                output_path = process_ots_data(
                    outstanding_il_file,
                    outstanding_jlg_file,
                    writeoff_il_file,
                    writeoff_jlg_file,
                    progress_callback=update_progress
                )

                with open(output_path, "rb") as f:
                    st.success("OTS Data Report generated successfully.")

                    st.download_button(
                        label="⬇️ Download OTS Data Report",
                        data=f,
                        file_name="OTS_Data_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error while generating OTS Data Report: {e}")

elif selected_report == "6) Single Client OD in a Center":
    st.subheader("6) Single Client OD in a Center")

    arrear_file = st.file_uploader(
        "Upload Arrear Report",
        type=["xlsx", "xls", "xlsb", "csv"],
        key="single_client_od_arrear"
    )

    if st.button("Generate Single Client OD Report", key="generate_single_client_od"):
        if arrear_file is None:
            st.error("Please upload Arrear Report.")
        else:
            with st.spinner("Processing Single Client OD Report..."):
                output_path = process_single_client_od(arrear_file)

            with open(output_path, "rb") as f:
                st.success("Single Client OD Report generated successfully.")
                st.download_button(
                    label="Download Single Client OD Report",
                    data=f,
                    file_name="Single_Client_OD_in_Center.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
elif selected_report == "7) OD Amount Less 1000 Non NPA Clients":
    st.subheader("7) OD Amount Less 1000 Non NPA Clients")

    arrear_file = st.file_uploader(
        "Upload Arrear Report",
        type=["xlsx", "xls", "xlsb", "csv"],
        key="od_amount_less_1000_arrear"
    )

    if st.button("Generate OD Amount Less 1000 Report", key="generate_od_amount_less_1000"):
        if arrear_file is None:
            st.error("Please upload Arrear Report.")
        else:
            with st.spinner("Processing OD Amount Less 1000 Report..."):
                output_path = process_od_amount_less_1000(arrear_file)

            with open(output_path, "rb") as f:
                st.success("OD Amount Less 1000 Report generated successfully.")
                st.download_button(
                    label="Download OD Amount Less 1000 Report",
                    data=f,
                    file_name="OD_Amount_Less_1000_Non_NPA_Clients.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
elif selected_report == "8) OD Status of Two Loans":
    st.subheader("OD Status of Two Loans")

    uploaded_file = st.file_uploader(
        "Upload OD Status input file",
        type=["xlsx", "xls", "xlsb", "csv"]
    )

    if uploaded_file and st.button("Generate Report"):
        try:
            progress = st.progress(0)
            status_text = st.empty()

            def update_progress(percent, message):
                progress.progress(percent)
                status_text.info(message)

            output_file = process_od_status_two_loans(
                uploaded_file=uploaded_file,
                output_dir="outputs",
                progress_callback=update_progress
            )

            with open(output_file, "rb") as f:
                st.success("OD Status of Two Loans Report generated successfully.")
                st.download_button(
                    label="Download Report",
                    data=f,
                    file_name="OD_Status_of_Two_Loans_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error: {e}")
elif selected_report == "9) Midfin and Finpage Collection Recon":

    st.subheader("Midfin and Finpage Collection Recon")

    col1, col2 = st.columns(2)

    with col1:
        finpage_jlg = st.file_uploader(
            "Upload Finpage JLG Repayment",
            type=["xlsx", "xls", "xlsb", "csv"],
            key="finpage_jlg"
        )

        midfin_regular = st.file_uploader(
            "Upload Midfin Regular Collection",
            type=["xlsx", "xls", "xlsb"],
            key="midfin_regular"
        )

        branch_master = st.file_uploader(
            "Upload Branch Master",
            type=["xlsx", "xls", "xlsb"],
            key="branch_master"
        )

    with col2:
        finpage_il = st.file_uploader(
            "Upload Finpage IL Repayment",
            type=["xlsx", "xls", "xlsb", "csv"],
            key="finpage_il"
        )

        midfin_od = st.file_uploader(
            "Upload Midfin OD Collection",
            type=["xlsx", "xls", "xlsb"],
            key="midfin_od"
        )

    if st.button("Generate Report"):
        if not all([finpage_jlg, finpage_il, midfin_regular, midfin_od, branch_master]):
            st.error("Please upload all 5 files.")
        else:
            with st.spinner("Generating report..."):
                output_path = process_midfin_finpage_recon(
                    finpage_jlg,
                    finpage_il,
                    midfin_regular,
                    midfin_od,
                    branch_master,
                    output_dir="outputs"
                )

            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Midfin and Finpage Collection Recon",
                    data=f,
                    file_name="Midfin_and_Finpage_Collection_Recon.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
