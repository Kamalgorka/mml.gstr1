import streamlit as st
from ui import load_global_css

# ✅ MUST BE FIRST
st.set_page_config(page_title="MML Smart Reports", layout="wide")

# ✅ AFTER that
load_global_css()

st.title("📊 MML Smart Reports")
st.write("Select a team from the left sidebar.")
