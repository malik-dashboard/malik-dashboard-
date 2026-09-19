import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="Malik - WIR Dashboard", layout="wide")
st.title("📊 Malik - WIR Inspection Dashboard - 55 Records")

# File Load
df = None
if os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")
up = st.file_uploader("Excel Upload (First time only)", type=["xlsx"])
if up:
    df = pd.read_excel(up)
if df is None:
    st.stop()

st.success(f"Total WIR Loaded: {len(df)}")

# --- Charts - No Matplotlib Needed - Streamlit Native ---
status_col = [c for c in df.columns if 'status' in c.lower()][0]
st.subheader("Status Breakdown")
st.bar_chart(df[status_col].value_counts())

origin_col = [c for c in df.columns if 'originator' in c.lower()][0]
st.subheader("Originator Wise")
st.bar_chart(df[origin_col].value_counts())

# --- CLEAN EXCEL WITH CHART DATA ---
def make_excel():
    output = io.BytesIO()
    # Sirf kaam ki columns
    keep = [c for c in df.columns if 'abcc' not in c.lower()][:12]
    clean = df[keep]
    summary = pd.DataFrame({
        'Status': df[status_col].value_counts().index,
        'Count': df[status_col].value_counts().values
    })
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        clean.to_excel(writer, index=False, sheet_name='WIR_Data')
        summary.to_excel(writer, index=False, sheet_name='Chart_Summary')
        # Originator summary for chart
        df[origin_col].value_counts().reset_index().to_excel(writer, index=False, sheet_name='Originator_Chart')
    return output.getvalue()

st.divider()
st.download_button(
    label="📊 CHART DATA KE SAATH EXCEL DOWNLOAD - CLICK HERE",
    data=make_excel(),
    file_name="Malik_WIR_With_Chart_Data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary"
)
