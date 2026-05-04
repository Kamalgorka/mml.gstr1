import streamlit as st
import pandas as pd
from auth_utils import load_users, save_users, log_activity
from ui import load_global_css

# MUST BE FIRST
st.set_page_config(page_title="MML Smart Reports", layout="wide")

load_global_css()

# ---------------- SESSION INIT ----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "must_change_password" not in st.session_state:
    st.session_state.must_change_password = False


# ---------------- LOGIN CSS ----------------
def load_login_css():
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none; }
    [data-testid="collapsedControl"] { display: none; }

    .block-container {
        max-width: 500px;
        margin: auto;
        padding-top: 80px;
    }

    .login-card {
        background: white;
        padding: 40px;
        border-radius: 18px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.12);
        border: 1px solid #e5e7eb;
    }

    .login-title {
        text-align: center;
        font-size: 30px;
        font-weight: 800;
        color: #1f4ed8;
        margin-bottom: 5px;
    }

    .login-subtitle {
        text-align: center;
        color: #6b7280;
        font-size: 14px;
        margin-bottom: 25px;
    }

    div.stButton > button {
        width: 100%;
        background: #1f4ed8;
        color: white;
        border-radius: 10px;
        height: 45px;
        font-weight: 700;
        border: none;
    }

    div.stButton > button:hover {
        background: #173bb0;
    }
    </style>
    """, unsafe_allow_html=True)


# ---------------- HIDE DEFAULT STREAMLIT PAGES ----------------
def hide_streamlit_pages():
    st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {
        display: none;
    }
    </style>
    """, unsafe_allow_html=True)


# ---------------- LOGIN FUNCTION ----------------
def login_user(email, password):
    df = load_users()

    user = df[(df["email"] == email) & (df["status"] == "ACTIVE")]

    if user.empty:
        return False, "Invalid user"

    user_row = user.iloc[0]

    if user_row["password"] != password:
        return False, "Wrong password"

    st.session_state.logged_in = True
    st.session_state.user_email = email
    st.session_state.user_role = user_row["role"].strip().upper()
    st.session_state.must_change_password = user_row["must_change_password"] == "YES"

    log_activity("LOGIN")

    return True, "Login successful"


# ---------------- PASSWORD CHANGE ----------------
def change_password(new_pass):
    df = load_users()

    df.loc[df["email"] == st.session_state.user_email, "password"] = new_pass
    df.loc[df["email"] == st.session_state.user_email, "must_change_password"] = "NO"

    save_users(df)

    st.session_state.must_change_password = False
    log_activity("PASSWORD_CHANGE")


# ---------------- LOGIN SCREEN ----------------
if not st.session_state.logged_in:

    load_login_css()

    st.markdown("""
    <div class="login-card">
        <div class="login-title">🔐 MML Smart Reports</div>
        <div class="login-subtitle">Please sign in to continue</div>
    </div>
    """, unsafe_allow_html=True)

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        success, msg = login_user(email.strip().lower(), password)

        if success:
            st.success("Login successful")
            st.rerun()
        else:
            st.error(msg)

    st.stop()


# ---------------- FORCE PASSWORD CHANGE ----------------
if st.session_state.must_change_password:

    load_login_css()

    st.markdown("""
    <div class="login-card">
        <div class="login-title">🔑 Change Password</div>
        <div class="login-subtitle">Password change is mandatory</div>
    </div>
    """, unsafe_allow_html=True)

    new_pass = st.text_input("New Password", type="password")
    confirm_pass = st.text_input("Confirm Password", type="password")

    if st.button("Update Password"):

        if not new_pass:
            st.error("Password cannot be empty")

        elif new_pass != confirm_pass:
            st.error("Passwords do not match")

        else:
            change_password(new_pass)
            st.success("Password updated successfully")
            st.rerun()

    st.stop()


# ---------------- MAIN APP ----------------

hide_streamlit_pages()

role = st.session_state.get("user_role", "")

st.sidebar.title("📌 Menu")

st.sidebar.page_link("streamlit_app.py", label="Home", icon="🏠")

# IMPORTANT: FILE NAMES FROM YOUR SCREENSHOT
if role in ["REPORTING", "ADMIN"]:
    st.sidebar.page_link("pages/2_Reporting_Team.py", label="Reporting Team", icon="📊")

if role in ["HO", "ADMIN"]:
    st.sidebar.page_link("pages/1_HO_Team.py", label="HO Team", icon="🏢")

st.title("Welcome to MML Smart Reports")
