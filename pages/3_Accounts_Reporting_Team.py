import streamlit as st

st.set_page_config(
    page_title="Accounts Reporting Team",
    page_icon="📘",
    layout="wide"
)

st.title("📘 Accounts Reporting Team - Reports Automisation")

st.markdown("Select report from below and upload required files to generate output.")

# Report selection
report_options = [
    "Select Report",
    "1) Sample Accounts Report"
]

selected_report = st.selectbox(
    "🔎 Select Report (type to search)",
    report_options
)

st.markdown("---")

if selected_report == "Select Report":
    st.info("Please select a report to continue.")

elif selected_report == "1) Sample Accounts Report":
    st.subheader("📘 Sample Accounts Report")

    st.caption("Upload required files > Run Automation > Download Output")

    file1 = st.file_uploader(
        "1) Upload Input File",
        type=["xlsx", "xls", "csv"]
    )

    run_btn = st.button("🚀 Run Automation")

    if run_btn:
        if file1 is None:
            st.error("Please upload the required file.")
        else:
            st.success("File uploaded successfully. Automation logic will be added here.")
