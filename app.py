import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>Any Excel File | Auto Charts</p>
</div>
""", unsafe_allow_html=True)

f = st.file_uploader("📂 Koi bhi Excel File Upload Karo", type=["xlsx","xls"])

if f:
    df = pd.read_excel(f)
elif os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")
else:
    st.info("👆 File upload karo")
    st.stop()

st.success(f"✅ Loaded: {len(df)} Records | {len(df.columns)} Columns")

# Safe metrics
c1,c2,c3,c4 = st.columns(4)
c1.metric("TOTAL ROWS", len(df))
c2.metric("COLUMNS", len(df.columns))
c3.metric("TEXT COLS", len(df.select_dtypes(include='object').columns))
c4.metric("NUMERIC", len(df.select_dtypes(include='number').columns))

st.divider()

# Only important columns for chart - Unwanted_ columns hatao
good_cols = [c for c in df.columns if 'Unnamed' not in str(c) and 'Source_File' not in str(c)]
text_cols = df[good_cols].select_dtypes(include=['object']).columns.tolist()

# Sirf 4 best columns lo - jo zyada useful hain
# Status, Originator jaisa kuch ho to pehle lo
priority = ['status','originator','stage','step','phase','type','discipline','area']
sorted_cols = []
for p in priority:
    for c in text_cols:
        if p in c.lower() and c not in sorted_cols:
            sorted_cols.append(c)
for c in text_cols:
    if c not in sorted_cols:
        sorted_cols.append(c)

chart_cols = sorted_cols[:4]

if len(chart_cols) >= 1:
    a,b = st.columns(2)
    with a:
        col = chart_cols[0]
        vc = df[col].astype(str).str.strip().value_counts().head(8)
        d = vc.reset_index()
        d.columns=[col,'Count']
        fig = px.pie(d, values='Count', names=col, hole=0.5, title=f"{col} - Top 8")
        fig.update_traces(textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
    if len(chart_cols) >= 2:
        with b:
            col = chart_cols[1]
            vc = df[col].astype(str).value_counts().head(10).reset_index()
            vc.columns=[col,'Count']
            fig = px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} - Top 10")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)

if len(chart_cols) >= 3:
    c,dcol = st.columns(2)
    with c:
        col = chart_cols[2]
        vc = df[col].astype(str).value_counts().head(10).reset_index()
        vc.columns=[col,'Count']
        fig = px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} - Top 10")
        fig.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig, use_container_width=True)
    if len(chart_cols) >= 4:
        with dcol:
            col = chart_cols[3]
            vc = df[col].astype(str).value_counts().head(10).reset_index()
            vc.columns=[col,'Count']
            fig = px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} - Top 10")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)

st.markdown("### 📋 Data Preview (First 30 Rows)")
st.dataframe(df[good_cols].head(30), use_container_width=True)

# FIXED DOWNLOAD - Error nahi ayega
def make_excel():
    out = io.BytesIO()
    try:
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            df[good_cols].to_excel(w, index=False, sheet_name='Data')
            for i, col in enumerate(chart_cols):
                safe_name = f"Chart{i+1}_{col}"[:31].replace(':','').replace('/','').replace('\\','')
                try:
                    df[col].value_counts().reset_index().to_excel(w, index=False, sheet_name=safe_name)
                except:
                    pass
    except Exception as e:
        st.error(f"Excel error: {e}")
        return None
    return out.getvalue()

excel_data = make_excel()
if excel_data:
    st.download_button(f"📥 DOWNLOAD - {len(df)} Records Excel", excel_data, file_name=f"Report_{len(df)}_Records.xlsx", use_container_width=True, type="primary")
