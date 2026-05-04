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

# ---------------- LOGIN FUNCTION ----------------
def login_user(email, password):
    df = load_users()

    user = df[(df["email"] == email) & (df["status"] == "ACTIVE")]

    if user.empty:
        return False, "Invalid user"

    user_row = user.iloc[0]

    if user_row["password"] != password:
        return False, "Wrong password"

    # Set session
    st.session_state.logged_in = True
    st.session_state.user_email = email
    st.session_state.user_role = user_row["role"]
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

    st.title("🔐 Login - MML Smart Reports")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        success, msg = login_user(email, password)

        if success:
            st.success("Login successful")
            st.rerun()
        else:
            st.error(msg)

    st.stop()

# ---------------- FORCE PASSWORD CHANGE ----------------
if st.session_state.must_change_password:

    st.title("⚠️ Change Password (Mandatory)")

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
st.title("📊 MML Smart Reports")
st.write(f"Logged in as: {st.session_state.user_email} ({st.session_state.user_role})")

st.sidebar.success("Select a team from sidebar")

# Logout
if st.sidebar.button("Logout"):
    log_activity("LOGOUT")
    st.session_state.clear()
    st.rerun()
