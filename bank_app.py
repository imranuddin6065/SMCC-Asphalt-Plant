import streamlit as st
import pandas as pd
import os
from fpdf import FPDF
import io
from datetime import datetime

# --- LOGIN SETTINGS ---
# Aap apna manpasand Username aur Password yahan set kar sakte hain
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

# --- AGAY KA CODE SIRF LOGIN KE BAAD CHALEGA ---

# Excel file setup
EXCEL_FILE = "smcc_ledger.xlsx"
COLS = ["S.No", "Date", "Particulars", "Dr", "Cr", "Balance"]

if not os.path.exists(EXCEL_FILE):
    df_new = pd.DataFrame(columns=COLS)
    df_new.to_excel(EXCEL_FILE, index=False)

st.set_page_config(page_title="SMCC Ledger", layout="wide")

# Logout Button in Sidebar
if st.sidebar.button("Log Out"):
    st.session_state["logged_in"] = False
    st.rerun()

# --- HEADER ---
st.markdown("<h1 style='text-align: center;'>SMCC Asphalt Plant</h1>", unsafe_allow_html=True)
h_col1, h_col2 = st.columns(2)
h_col1.markdown("### **Meezan Account**")
h_col2.markdown("<h3 style='text-align: right;'><b>Imran Shah</b></h3>", unsafe_allow_html=True)
st.divider()

df = pd.read_excel(EXCEL_FILE)

# --- ENTRY FORM ---
with st.expander("➕ Nayi Entry Dalein"):
    with st.form("entry_form", clear_on_submit=True):
        selected_date = st.date_input("Tareekh Select Karein")
        desc = st.text_input("Particulars (Tafseel)")
        ttype = st.selectbox("Type", ["Cash Out (Dr)", "Cash In (Cr)"])
        amount = st.number_input("Raqam", min_value=0)
        submit = st.form_submit_button("Save Karein")

if submit:
    last_balance = df["Balance"].iloc[-1] if not df.empty else 0
    dr_val = amount if ttype == "Cash Out (Dr)" else 0
    cr_val = amount if ttype == "Cash In (Cr)" else 0
    new_balance = last_balance + cr_val - dr_val
    formatted_date = selected_date.strftime("%d-%b-%y")
    
    new_row = {"S.No": len(df) + 1, "Date": formatted_date, "Particulars": desc, "Dr": dr_val, "Cr": cr_val, "Balance": new_balance}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    st.success(f"Record Save Ho Gaya!")
    st.rerun()

# --- TABLE ---
st.subheader("Ledger Details")
st.dataframe(df, use_container_width=True, hide_index=True)

# --- DOWNLOADS ---
st.divider()
st.subheader("📥 Reports Download")
d_col1, d_col2 = st.columns(2)

# Excel Logic
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Ledger')
    workbook  = writer.book
    worksheet = writer.sheets['Ledger']
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#FFFF00', 'border': 1})
    cell_fmt = workbook.add_format({'align': 'center', 'border': 1})
    num_fmt = workbook.add_format({'align': 'center', 'border': 1, 'num_format': '#,##0'})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    worksheet.set_column('A:A', 8, cell_fmt)
    worksheet.set_column('B:B', 15, cell_fmt)
    worksheet.set_column('C:C', 45, cell_fmt)
    worksheet.set_column('D:F', 18, num_fmt)
d_col1.download_button("Excel Download", data=buffer.getvalue(), file_name="SMCC_Ledger.xlsx")

# PDF Logic
def generate_pdf(data):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "SMCC Asphalt Plant", ln=True, align='C')
    pdf.set_font("Arial", 'B', 11)
    y_pos = pdf.get_y()
    pdf.cell(0, 10, "Meezan Account", align='L')
    pdf.set_y(y_pos)
    pdf.cell(0, 10, "Imran Shah", align='R')
    pdf.ln(15)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(240, 240, 240)
    col_widths = [12, 28, 65, 25, 25, 30] 
    headers = ["S.No", "Date", "Particulars", "Dr", "Cr", "Balance"]
    start_x = (210 - sum(col_widths)) / 2
    pdf.set_x(start_x)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 10, h, 1, 0, 'C', fill=True)
    pdf.ln()
    pdf.set_font("Arial", '', 9)
    for _, row in data.iterrows():
        pdf.set_x(start_x)
        pdf.cell(col_widths[0], 10, str(row['S.No']), 1, 0, 'C')
        pdf.cell(col_widths[1], 10, str(row['Date']), 1, 0, 'C')
        pdf.cell(col_widths[2], 10, " " + str(row['Particulars'])[:40], 1, 0, 'L')
        pdf.cell(col_widths[3], 10, f"{row['Dr']:,.0f}", 1, 0, 'C')
        pdf.cell(col_widths[4], 10, f"{row['Cr']:,.0f}", 1, 0, 'C')
        pdf.cell(col_widths[5], 10, f"{row['Balance']:,.0f}", 1, 0, 'C')
        pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

if not df.empty:
    pdf_bytes = generate_pdf(df)
    d_col2.download_button("PDF Download", data=pdf_bytes, file_name="SMCC_Report.pdf")