import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>Auto Header Detection | Any Excel File</p>
</div>
""", unsafe_allow_html=True)

def smart_read(file):
    # Try to find real header row
    try:
        tmp = pd.read_excel(file, header=None)
        header_row = 0
        for i in range(min(10, len(tmp))):
            row = tmp.iloc[i].astype(str).str.lower()
            # Jahan Status, WIR, Originator jaisa lafz mile wahi header hai
            if any(x in ' '.join(row) for x in ['status','wir','origin','stage','discipline']):
                header_row = i
                break
        file.seek(0)
        df = pd.read_excel(file, header=header_row)
        df = df.dropna(how='all').dropna(axis=1, how='all')
        # Unnamed columns hatao
        df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed', na=False)]
        return df
    except Exception as e:
        file.seek(0)
        return pd.read_excel(file)

f = st.file_uploader("📂 Koi bhi Excel File Upload Karo", type=["xlsx","xls"])

if f:
    df = smart_read(f)
elif os.path.exists("data.xlsx"):
    df = smart_read(open("data.xlsx","rb"))
else:
    st.info("👆 File upload karo")
    st.stop()

st.success(f"✅ Loaded: {len(df)} Records | {len(df.columns)} Columns")
st.write("Columns Found:", list(df.columns)[:10])

c1,c2,c3,c4 = st.columns(4)
c1.metric("ROWS", len(df))
c2.metric("COLS", len(df.columns))
c3.metric("TEXT", len(df.select_dtypes(include='object').columns))
c4.metric("NUM", len(df.select_dtypes(include='number').columns))

# Clean for charts
good_cols = [c for c in df.columns if str(c).strip()!= '']
text_cols = df[good_cols].select_dtypes(include=['object']).columns.tolist()

# Priority
priority = ['status','originator','stage','discipline','type','area','system']
sorted_cols = []
for p in priority:
    for c in text_cols:
        if p in str(c).lower() and c not in sorted_cols:
            sorted_cols.append(c)
for c in text_cols:
    if c not in sorted_cols:
        sorted_cols.append(c)
chart_cols = sorted_cols[:4]

st.divider()

if chart_cols:
    a,b = st.columns(2)
    for idx, col in enumerate(chart_cols[:2]):
        vc = df[col].astype(str).str.strip().replace('nan','').replace('','')
        vc = vc[vc!=''].value_counts().head(10).reset_index()
        if len(vc)==0: continue
        vc.columns=[col,'Count']
        fig = px.pie(vc, values='Count', names=col, hole=0.5, title=f"{col} - Breakdown") if idx==0 else px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} - Top 10")
        if idx==0:
            fig.update_traces(textinfo='percent+label')
        else:
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
        (a if idx==0 else b).plotly_chart(fig, use_container_width=True)

    if len(chart_cols) > 2:
        c,d = st.columns(2)
        for idx, col in enumerate(chart_cols[2:4]):
            vc = df[col].astype(str).str.strip().replace('nan','').replace('','')
            vc = vc[vc!=''].value_counts().head(10).reset_index()
            if len(vc)==0: continue
            vc.columns=[col,'Count']
            fig = px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} - Top 10")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            (c if idx==0 else d).plotly_chart(fig, use_container_width=True)

st.markdown("### 📋 Data Preview")
st.dataframe(df.head(30), use_container_width=True)

# Download fixed
out = io.BytesIO()
with pd.ExcelWriter(out, engine='openpyxl') as w:
    df.to_excel(w, index=False, sheet_name='Data')
st.download_button(f"📥 DOWNLOAD - {len(df)} Records", out.getvalue(), file_name=f"Report_{len(df)}.xlsx", use_container_width=True, type="primary")
