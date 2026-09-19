import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0; opacity:0.9'>Any Excel File | Auto Charts | Professional Report</p>
</div>
""", unsafe_allow_html=True)

# --- Koi bhi file ---
f = st.file_uploader("📂 Koi bhi Excel File Upload Karo - WIR, LOD, Material, Koi bhi", type=["xlsx","xls"])

if f:
    df = pd.read_excel(f)
elif os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")
else:
    st.info("👆 Koi bhi Excel file upload karo")
    st.stop()

st.success(f"✅ File Loaded: {len(df)} Records | Columns: {len(df.columns)}")

# Metrics
c1,c2,c3,c4 = st.columns(4)
c1.metric("TOTAL ROWS", len(df))
c2.metric("TOTAL COLUMNS", len(df.columns))
c3.metric("NUMERIC COLS", len(df.select_dtypes(include='number').columns))
c4.metric("TEXT COLS", len(df.select_dtypes(include='object').columns))

st.divider()

# Auto Charts for ANY data - pehle 4 text columns ke charts
text_cols = df.select_dtypes(include=['object']).columns.tolist()[:4]

if len(text_cols) >= 1:
    col1, col2 = st.columns(2)
    with col1:
        d = df[text_cols[0]].astype(str).value_counts().head(10).reset_index()
        d.columns=[text_cols[0], 'Count']
        fig = px.pie(d, values='Count', names=text_cols[0], hole=0.5, title=f"{text_cols[0]} Breakdown")
        st.plotly_chart(fig, use_container_width=True)
    if len(text_cols) >= 2:
        with col2:
            d = df[text_cols[1]].astype(str).value_counts().head(15).reset_index()
            d.columns=[text_cols[1], 'Count']
            fig = px.bar(d, x=text_cols[1], y='Count', text='Count', title=f"{text_cols[1]} Wise", color='Count')
            st.plotly_chart(fig, use_container_width=True)

if len(text_cols) >= 3:
    col3, col4 = st.columns(2)
    with col3:
        d = df[text_cols[2]].astype(str).value_counts().head(15).reset_index()
        d.columns=[text_cols[2], 'Count']
        fig = px.bar(d, x=text_cols[2], y='Count', text='Count', title=f"{text_cols[2]} Wise", color='Count')
        st.plotly_chart(fig, use_container_width=True)
    if len(text_cols) >= 4:
        with col4:
            d = df[text_cols[3]].astype(str).value_counts().head(15).reset_index()
            d.columns=[text_cols[3], 'Count']
            fig = px.bar(d, x=text_cols[3], y='Count', text='Count', title=f"{text_cols[3]} Wise", color='Count')
            st.plotly_chart(fig, use_container_width=True)

# Data preview
st.markdown("### 📋 Data Preview")
st.dataframe(df.head(50), use_container_width=True)

# Download - Any file
out = io.BytesIO()
with pd.ExcelWriter(out, engine='openpyxl') as w:
    df.to_excel(w, index=False, sheet_name='Original_Data')
    for col in text_cols:
        df[col].value_counts().reset_index().to_excel(w, index=False, sheet_name=f'Chart_{col}'[:31])

st.download_button(f"📥 DOWNLOAD REPORT - {len(df)} Records", out.getvalue(), file_name=f"Report_{len(df)}_Records.xlsx", use_container_width=True, type="primary")
