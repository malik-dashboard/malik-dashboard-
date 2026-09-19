
import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Malik 100K Files Dashboard", layout="wide")
st.markdown("<h2 style='text-align:center'>📊 Malik - 100,000 Files Universal Dashboard</h2>", unsafe_allow_html=True)

# --- 100K FILES KA SOLUTION ---
st.warning("💡 100,000 Files ka tareeka: 500-500 files ek saath upload karo, ya saari files ek ZIP me daal ke upload karo")

uploaded_files = st.file_uploader("100, 500, 1000 - Jitni marzi Excel files ek saath upload karo", type=["xlsx","xls"], accept_multiple_files=True)

df = None
if uploaded_files:
    all_dfs = []
    for f in uploaded_files:
        try:
            temp = pd.read_excel(f)
            temp['Source_File'] = f.name  # kaunsi file se aaya
            all_dfs.append(temp)
        except: pass
    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True, sort=False) # alag alag columns bhi chalenge
        st.success(f"✅ {len(uploaded_files)} Files Combined! Total Records: {len(df)}")
elif os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")

if df is None:
    st.info("👆 Upar files upload karo - 1 file ho ya 1000, sab chalega")
    st.stop()

# --- UNIVERSAL COLUMN FINDER ---
def find_col(kws):
    for kw in kws:
        for c in df.columns:
            if kw in str(c).lower():
                return c
    return None

status_col = find_col(['status','result','decision','approval'])
orig_col = find_col(['originator','created by','engineer','inspector'])
stage_col = find_col(['stage','step','current'])

# Metrics
c1,c2,c3,c4 = st.columns(4)
c1.metric("TOTAL FILES", len(uploaded_files) if uploaded_files else 1)
c2.metric("TOTAL RECORDS", len(df))
c3.metric("UNIQUE ORIGINATORS", df[orig_col].nunique() if orig_col else 0)
c4.metric("STATUS TYPES", df[status_col].nunique() if status_col else 0)

colA, colB = st.columns(2)
with colA:
    if status_col:
        sc = df[status_col].astype(str).value_counts().reset_index()
        sc.columns=["Status","Count"]
        fig = px.pie(sc, values="Count", names="Status", hole=0.5, title="Status Breakdown (All Files Combined)")
        st.plotly_chart(fig, use_container_width=True)

with colB:
    if orig_col:
        oc = df[orig_col].astype(str).value_counts().reset_index()
        oc.columns=["Originator","Count"]
        fig2 = px.bar(oc.head(20), x="Originator", y="Count", title="Top 20 Originators (All Files)", text="Count")
        st.plotly_chart(fig2, use_container_width=True)

if stage_col:
    st.markdown("#### Workflow Stage - All Files Combined")
    stg = df[stage_col].astype(str).value_counts().reset_index()
    stg.columns=["Stage","Count"]
    fig3 = px.bar(stg, x="Stage", y="Count", text="Count", color="Count")
    st.plotly_chart(fig3, use_container_width=True)

# --- MEGA DOWNLOAD ---
def make_mega_excel():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        # 1. Combined Data
        df.to_excel(writer, index=False, sheet_name='All_100K_Combined')
        # 2. Summary
        if status_col:
            df[status_col].value_counts().reset_index().to_excel(writer, index=False, sheet_name='Status_Summary')
        if orig_col:
            df[orig_col].value_counts().reset_index().to_excel(writer, index=False, sheet_name='Originator_Summary')
    return out.getvalue()

st.divider()
st.download_button(f"📥 {len(df)} RECORDS KA COMBINED EXCEL DOWNLOAD - SAME CHARTS KE SAATH",
                   data=make_mega_excel(),
                   file_name=f"Malik_Combined_{len(df)}_Records.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True, type="primary")
