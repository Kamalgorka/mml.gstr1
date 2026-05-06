import os
import csv
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.csv")
LOG_FILE = os.path.join(BASE_DIR, "activity_log.csv")


def load_users():
    if not os.path.exists(USERS_FILE):
        st.error("users.csv file not found on server.")
        st.stop()
    return pd.read_csv(USERS_FILE, dtype=str).fillna("")


def save_users(df):
    df.to_csv(USERS_FILE, index=False)


def log_activity(activity, report="", status="SUCCESS"):
    file_exists = os.path.exists(LOG_FILE)

    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        if not file_exists:
            writer.writerow(["datetime", "email", "role", "activity", "report", "status"])

        writer.writerow([
            datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.get("user_email", ""),
            st.session_state.get("user_role", ""),
            activity,
            report,
            status
        ])


def require_login():
    if not st.session_state.get("logged_in"):
        st.warning("Please login first.")
        st.stop()


def require_role(allowed_roles):
    import streamlit as st

    if "logged_in" not in st.session_state or not st.session_state.logged_in:
        st.warning("🔐 Please login first")
        st.switch_page("streamlit_app.py")
        st.stop()

    user_role = st.session_state.get("user_role", "").strip().upper()

    allowed_roles = [r.strip().upper() for r in allowed_roles]

    if user_role not in allowed_roles:
        st.error("⛔ Access Denied")
        st.switch_page("streamlit_app.py")
        st.stop()


def run_with_logging(report_name, func):
    try:
        log_activity("RUN_REPORT", report_name, "STARTED")
        result = func()
        log_activity("RUN_REPORT", report_name, "SUCCESS")
        return result
    except Exception:
        log_activity("RUN_REPORT", report_name, "FAILED")
        raise
