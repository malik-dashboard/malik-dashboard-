import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>Any Excel File | Fixed Download</p>
</div>
""", unsafe_allow_html=True)

f = st.file_uploader("📂 Koi bhi Excel File Upload Karo", type=["xlsx","xls"])

if f:
    df = pd.read_excel(f, header=None, nrows=1)
    f.seek(0)
    xls = pd.ExcelFile(f)
    sheet = st.selectbox("Sheet Select Karo", xls.sheet_names)
    raw = pd.read_excel(f, sheet_name=sheet, header=None, nrows=8)
    st.dataframe(raw, use_container_width=True)
    header_row = st.number_input("Header Row Number? (Jahan Status likha hai)", 0, 7, 5)
    f.seek(0)
    df = pd.read_excel(f, sheet_name=sheet, header=header_row)
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False, regex=True)]
else:
    if os.path.exists("data.xlsx"):
        df = pd.read_excel("data.xlsx")
    else:
        st.info("👆 File upload karo")
        st.stop()

st.success(f"✅ Loaded: {len(df)} Records")

# --- FIXED DOWNLOAD - YEH LINE 48 SE 62 TAK HAI - YAHIN FIX HAI ---
def get_excel_file():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    output.seek(0)
    return output.getvalue()

st.dataframe(df.head(30), use_container_width=True)

st.download_button(
    label=f"📥 DOWNLOAD FIXED EXCEL - {len(df)} Records",
    data=get_excel_file(),
    file_name=f"Report_Fixed_{len(df)}_Records.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True, 
    type="primary"
)
