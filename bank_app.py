import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# --- 1. GOOGLE SHEETS SETUP ---
def get_gspread_client():
    # Streamlit Secrets se credentials uthana
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds_info = st.secrets["gcp_service_account"]
        credentials = Credentials.from_service_account_info(creds_info, scopes=scope)
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"Secrets mein masla hai: {e}")
        return None

# Sheet ka naam (Jo aapne Google Sheets mein rakha hai)
SHEET_NAME = "SMCC_Ledger_Data"

def load_data():
    client = get_gspread_client()
    if client:
        try:
            sh = client.open(SHEET_NAME)
            worksheet = sh.get_worksheet(0)
            data = worksheet.get_all_records()
            return pd.DataFrame(data)
        except Exception as e:
            st.error(f"Sheet nahi mil rahi: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def append_data(row_list):
    client = get_gspread_client()
    if client:
        sh = client.open(SHEET_NAME)
        worksheet = sh.get_worksheet(0)
        worksheet.append_row(row_list)

# --- 2. LOGIN SETTINGS ---
USER_CREDENTIALS = {"admin": "smcc123"} 

def login():
    st.markdown("<h2 style='text-align: center;'>SMCC Ledger Login</h2>", unsafe_allow_html=True)
    with st.form("login_form"):
        user = st.text_input("Username")
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if user in USER_CREDENTIALS and USER_CREDENTIALS[user] == pw:
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("Ghalat Username ya Password!")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login()
    st.stop()

# --- 3. APP INTERFACE ---
st.set_page_config(page_title="SMCC Ledger", layout="wide")

if st.sidebar.button("Log Out"):
    st.session_state["logged_in"] = False
    st.rerun()

st.markdown("<h1 style='text-align: center;'>SMCC Asphalt Plant</h1>", unsafe_allow_html=True)
h_col1, h_col2 = st.columns(2)
h_col1.markdown("### **Meezan Account**")
h_col2.markdown("<h3 style='text-align: right;'><b>Imran Shah</b></h3>", unsafe_allow_html=True)
st.divider()

# Data load karein
df = load_data()

if df.empty:
    st.warning("Google Sheet khali hai ya connect nahi hui. Pehli row mein headings (S.No, Date, Particular, Dr, Cr, Balance) likhein.")

# --- ENTRY FORM ---
with st.expander("➕ Nayi Entry Dalein"):
    with st.form("entry_form", clear_on_submit=True):
        selected_date = st.date_input("Tareekh Select Karein")
        desc = st.text_input("Particulars (Tafseel)")
        ttype = st.selectbox("Type", ["Cash Out (Dr)", "Cash In (Cr)"])
        amount = st.number_input("Raqam", min_value=0)
        submit = st.form_submit_button("Save Karein")

if submit:
    # Balance calculation
    last_balance = 0
    if not df.empty and "Balance" in df.columns:
        last_balance = df["Balance"].iloc[-1]
    
    dr_val = amount if ttype == "Cash Out (Dr)" else 0
    cr_val = amount if ttype == "Cash In (Cr)" else 0
    new_balance = float(last_balance) + float(cr_val) - float(dr_val)
    formatted_date = selected_date.strftime("%d-%b-%y")
    s_no = len(df) + 1
    
    # Google Sheet mein save karein
    new_row = [s_no, formatted_date, desc, dr_val, cr_val, new_balance]
    append_data(new_row)
    
    st.success("Record Save Ho Gaya!")
    st.rerun()

# --- TABLE ---
st.subheader("Live Ledger (Google Sheets)")
st.dataframe(df, use_container_width=True, hide_index=True)

# --- REPORTS ---
if not df.empty:
    st.divider()
    st.subheader("📥 Reports Download")
    d_col1, d_col2 = st.columns(2)

    # Excel Download
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Ledger')
    d_col1.download_button("Excel Download", data=buffer.getvalue(), file_name="SMCC_Ledger.xlsx")

    # PDF Download
    def generate_pdf(data):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(200, 10, "SMCC Ledger Report", ln=True, align='C')
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 10)
        for col in data.columns:
            pdf.cell(32, 10, str(col), 1)
        pdf.ln()
        pdf.set_font("Arial", '', 9)
        for _, row in data.iterrows():
            for val in row:
                pdf.cell(32, 10, str(val), 1)
            pdf.ln()
        return pdf.output(dest='S').encode('latin-1')

    pdf_bytes = generate_pdf(df)
    d_col2.download_button("PDF Download", data=pdf_bytes, file_name="SMCC_Report.pdf")
