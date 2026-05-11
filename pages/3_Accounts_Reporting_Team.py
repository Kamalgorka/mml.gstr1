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
    "1) SMS Report"
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

        hub56_file = st.file_uploader(
            "Monthly Outstanding SMS Data JLG HUB 5 and 6",
            type=["xlsx", "xls"],
            key="hub56"
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
        use_container_width=True
    )

    if run_btn:

        required_files = {
            "JLG HUB 1": hub1_file,
            "JLG HUB 2": hub2_file,
            "JLG HUB 3": hub3_file,
            "JLG HUB 4": hub4_file,
            "JLG HUB 5 and 6": hub56_file,
            "IL": il_file,
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

            from accounts_reports.sms_report import process_sms_report

            with st.spinner("Processing SMS Reports..."):
                result = process_sms_report(required_files)

            st.success("SMS raw files loaded successfully.")

            for k in result.keys():
                st.write(f"✅ {k}")
