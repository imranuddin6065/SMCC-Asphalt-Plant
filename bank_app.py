import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
from gsheetsdb import connect # Google Sheets connect karne ke liye
import gspread # Data likhne ke liye
from google.oauth2.service_account import Credentials

# --- 1. GOOGLE SHEETS SETUP ---
# Ye hissa aapke Streamlit Secrets se malomat uthaye ga
def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # Secrets se credentials lena
    creds_info = st.secrets["gcp_service_account"]
    credentials = Credentials.from_service_account_info(creds_info, scopes=scope)
    return gspread.authorize(credentials)

# Sheet ka naam (Jo aapne rakha tha)
SHEET_NAME = "SMCC_Ledger_Data"

def load_data():
    client = get_gspread_client()
    sh = client.open(SHEET_NAME)
    worksheet = sh.get_worksheet(0)
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

def append_data(row_list):
    client = get_gspread_client()
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

# Google Sheet se data load karein
try:
    df = load_data()
except Exception as e:
    st.error("Google Sheet connect nahi ho saki. Check karein ke aapne Sheet Share ki hai?")
    st.stop()

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
    last_balance = df["Balance"].iloc[-1] if not df.empty else 0
    dr_val = amount if ttype == "Cash Out (Dr)" else 0
    cr_val = amount if ttype == "Cash In (Cr)" else 0
    new_balance = float(last_balance) + float(cr_val) - float(dr_val)
    formatted_date = selected_date.strftime("%d-%b-%y")
    s_no = len(df) + 1
    
    # Google Sheet mein likhna (S.No, Date, Particular, Dr, Cr, Balance)
    new_row_list = [s_no, formatted_date, desc, dr_val, cr_val, new_balance]
    append_data(new_row_list)
    
    st.success(f"Record Google Sheet mein Save Ho Gaya!")
    st.rerun()

# --- TABLE ---
st.subheader("Live Ledger (Google Sheets)")
st.dataframe(df, use_container_width=True, hide_index=True)

# --- REPORTS ---
st.divider()
st.subheader("📥 Reports Download")
d_col1, d_col2 = st.columns(2)

# Excel Download
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Ledger')
d_col1.download_button("Excel Download", data=buffer.getvalue(), file_name="SMCC_Ledger.xlsx")

# PDF Download Logic
def generate_pdf(data):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "SMCC Asphalt Plant", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 10)
    col_widths = [12, 28, 65, 25, 25, 30] 
    headers = ["S.No", "Date", "Particulars", "Dr", "Cr", "Balance"]
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 10, h, 1, 0, 'C')
    pdf.ln()
    pdf.set_font("Arial", '', 9)
    for _, row in data.iterrows():
        pdf.cell(col_widths[0], 10, str(row['S.No']), 1, 0, 'C')
        pdf.cell(col_widths[1], 10, str(row['Date']), 1, 0, 'C')
        pdf.cell(col_widths[2], 10, str(row['Particulars'])[:35], 1, 0, 'L')
        pdf.cell(col_widths[3], 10, str(row['Dr']), 1, 0, 'C')
        pdf.cell(col_widths[4], 10, str(row['Cr']), 1, 0, 'C')
        pdf.cell(col_widths[5], 10, str(row['Balance']), 1, 0, 'C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

if not df.empty:
    try:
        pdf_bytes = generate_pdf(df)
        d_col2.download_button("PDF Download", data=pdf_bytes, file_name="SMCC_Report.pdf")
    except:
        st.info("PDF generate karne ke liye data check karein.")
