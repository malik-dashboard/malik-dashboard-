import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>Fix - Select Header Row</p>
</div>
""", unsafe_allow_html=True)

f = st.file_uploader("📂 WIR File Upload Karo", type=["xlsx","xls"])

if not f and os.path.exists("data.xlsx"):
    f = open("data.xlsx","rb")

if not f:
    st.info("File upload karo")
    st.stop()

# --- Step 1: Raw file dikhao ---
raw = pd.read_excel(f, header=None, nrows=8)
f.seek(0)
st.markdown("### ⚙️ Step 1: Header Row Select Karo - Yahan se Status milega!")
st.write("Neeche dekho - Jahan Status, WIR, Originator likha hai - Wo row number select karo:")
st.dataframe(raw, use_container_width=True)

header_row = st.number_input("Header Row Number Select Karo (0=Peheli line, 1=Dusri line, 2=Teesri line)", min_value=0, max_value=7, value=2)

# --- Step 2: Sahi header se load karo ---
f.seek(0)
df = pd.read_excel(f, header=header_row)
df = df.dropna(how='all').dropna(axis=1, how='all')
# Unnamed hatao
df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False, regex=True)]

st.success(f"✅ Loaded: {len(df)} Records | Columns: {list(df.columns)}")

# Ab Status dhoondo
def find_col(kws):
    for kw in kws:
        for c in df.columns:
            if kw in str(c).lower():
                return c
    return None

status_c = find_col(['status','result'])
orig_c = find_col(['originator','engineer','created by','raised'])
stage_c = find_col(['stage','step','phase'])

st.markdown(f"**Found Columns -> Status: {status_c} | Originator: {orig_c} | Stage: {stage_c}**")

c1,c2,c3,c4 = st.columns(4)
c1.metric("TOTAL", len(df))
c2.metric("STATUS TYPES", df[status_c].nunique() if status_c else 0)
c3.metric("ORIGINATORS", df[orig_c].nunique() if orig_c else 0)
c4.metric("STAGES", df[stage_c].nunique() if stage_c else 0)

st.divider()

# Charts - Ab Status se banega
if status_c:
    a,b = st.columns(2)
    with a:
        vc = df[status_c].astype(str).value_counts().reset_index()
        vc.columns=[status_c,'Count']
        fig = px.pie(vc, values='Count', names=status_c, hole=0.5, title=f"{status_c} Breakdown")
        fig.update_traces(textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
    if orig_c:
        with b:
            vc = df[orig_c].astype(str).value_counts().head(10).reset_index()
            vc.columns=[orig_c,'Count']
            fig = px.bar(vc, x='Count', y=orig_c, orientation='h', text='Count', title=f"{orig_c} Top 10")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)

st.dataframe(df.head(30), use_container_width=True)

out = io.BytesIO()
with pd.ExcelWriter(out, engine='openpyxl') as w:
    df.to_excel(w, index=False, sheet_name='Data')
st.download_button(f"📥 DOWNLOAD {len(df)} Records", out.getvalue(), file_name=f"Report_{len(df)}.xlsx", use_container_width=True, type="primary")
