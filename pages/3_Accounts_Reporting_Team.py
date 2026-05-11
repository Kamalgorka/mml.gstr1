import streamlit as st
from auth_utils import require_role, log_activity

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
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Accounts Reporting Team",
    page_icon="📘",
    layout="wide"
)

st.title("📘 Accounts Reporting Team - Reports Automisation")

st.markdown(
    "Select report from below and upload required files to generate output."
)

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

    st.caption(
        "Upload required files > Run Automation > Download Output"
    )

    file1 = st.file_uploader(
        "1) Upload Input File",
        type=["xlsx", "xls", "csv"]
    )

    run_btn = st.button("🚀 Run Automation")

    if run_btn:

        if file1 is None:
            st.error("Please upload the required file.")

        else:
            st.success(
                "File uploaded successfully. Automation logic will be added here."
            )
