import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>WIR Log Analytics | Professional</p>
</div>
""", unsafe_allow_html=True)

f = st.file_uploader("📂 Koi bhi Excel File Upload Karo", type=["xlsx","xls"])

if f:
    xls = pd.ExcelFile(f)
    sheet = st.selectbox("Sheet Select Karo", xls.sheet_names)
    # Header row selector - aapke screenshot me 5 sahi hai
    header_row = st.number_input("Header Row Number? (Screenshot me 5 sahi hai)", 0, 10, 5)
    f.seek(0)
    df = pd.read_excel(f, sheet_name=sheet, header=header_row)
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False, regex=True)]
elif os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")
else:
    st.info("👆 File upload karo")
    st.stop()

# Ab aapke screenshot wala data
if len(df) == 0:
    st.error("Header Row galat hai - 4 ya 5 try karo")
    st.stop()

st.success(f"✅ Loaded: {len(df)} Records | {len(df.columns)} Columns")

# --- METRICS ---
c1,c2,c3,c4 = st.columns(4)
c1.metric("TOTAL", len(df))
status_col = None
for c in df.columns:
    if 'Approved' in str(c) or 'Status' in str(c):
        status_col = c
        break
# Grand Total nikal lo
grand_col = None
for c in df.columns:
    if 'Grand Total' in str(c):
        grand_col = c
        break

if grand_col:
    try:
        c2.metric("TOTAL WIR", int(pd.to_numeric(df[grand_col], errors='coerce').sum()))
    except:
        c2.metric("COLUMNS", len(df.columns))
else:
    c2.metric("COLUMNS", len(df.columns))

c3.metric("SYSTEMS", df['Row Labels'].nunique() if 'Row Labels' in df.columns else 0)
c4.metric("STATUS TYPES", 3)

st.divider()

# --- CHARTS - Aapke pivot ke hisab se ---
if 'Row Labels' in df.columns:
    a,b = st.columns(2)
    with a:
        # Top 10 System by Grand Total
        if grand_col:
            plot_df = df[['Row Labels', grand_col]].copy()
            plot_df[grand_col] = pd.to_numeric(plot_df[grand_col], errors='coerce')
            plot_df = plot_df.sort_values(grand_col, ascending=False).head(10)
            fig = px.bar(plot_df, x=grand_col, y='Row Labels', orientation='h', text=grand_col, title="Top 10 Systems by WIR Count")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
    with b:
        # Status breakdown
        app_cols = [c for c in df.columns if 'A - Approved' in str(c) or 'A-Approved' in str(c)]
        b_cols = [c for c in df.columns if 'B - Approved' in str(c) or 'B-Approved' in str(c)]
        if app_cols and b_cols:
            total_a = pd.to_numeric(df[app_cols[0]], errors='coerce').sum()
            total_b = pd.to_numeric(df[b_cols[0]], errors='coerce').sum()
            pie_df = pd.DataFrame({'Status':['A-Approved','B-Approved with Comments'], 'Count':[total_a, total_b]})
            fig2 = px.pie(pie_df, values='Count', names='Status', hole=0.5, title="Overall Status Breakdown")
            st.plotly_chart(fig2, use_container_width=True)

st.markdown("### 📋 Data Preview")
st.dataframe(df.head(30), use_container_width=True)

# --- FIXED DOWNLOAD - CORRUPT ERROR FIXED ---
def get_excel_download():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='WIR_Data')
    output.seek(0)
    return output.getvalue()

st.download_button(
    label=f"📥 DOWNLOAD FIXED EXCEL - {len(df)} Records",
    data=get_excel_download(),
    file_name=f"WIR_Report_Fixed_{len(df)}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary"
)
