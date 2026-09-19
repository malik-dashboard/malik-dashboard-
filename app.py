import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>WIR + Any Excel File - Auto Detect</p>
</div>
""", unsafe_allow_html=True)

f = st.file_uploader("📂 Koi bhi Excel File Upload Karo (WIR, Material, Any)", type=["xlsx","xls"])

if f:
    xls = pd.ExcelFile(f)
    sheet = st.selectbox(f"Sheets Found: {len(xls.sheet_names)}", xls.sheet_names)

    raw = pd.read_excel(f, sheet_name=sheet, header=None, nrows=6)
    st.write("File ka Top View - Header kahan hai check karo:")
    st.dataframe(raw, use_container_width=True)

    header_row = st.number_input("Header Row Number (0,1,2,3,4,5 try karo)", 0, 10, 0)
    f.seek(0)
    df = pd.read_excel(f, sheet_name=sheet, header=header_row)
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False, regex=True)]
elif os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")
else:
    st.info("File upload karo - WIR ho ya koi bhi Excel")
    st.stop()

if len(df) == 0:
    st.error("Header Row galat hai - Dusra number try karo")
    st.stop()

st.success(f"✅ Loaded: {len(df)} Records | {len(df.columns)} Columns")

# --- AUTO DETECT: WIR PIVOT HAI YA NORMAL EXCEL? ---
is_pivot = 'Row Labels' in df.columns or 'Count of STATUS' in str(df.columns)

if is_pivot:
    st.warning("📌 WIR Pivot File Detect Hui - Status Wise Analysis")
    # WIR LOGIC
    df_count = df[[c for c in df.columns if '.1' not in str(c)]]
    grand_col = [c for c in df_count.columns if 'Grand Total' in str(c)][0] if any('Grand Total' in str(c) for c in df_count.columns) else None
    row_col = [c for c in df_count.columns if 'Row Labels' in str(c)][0] if any('Row Labels' in str(c) for c in df_count.columns) else df_count.columns[0]

    if grand_col:
        grand_idx = df_count.columns.get_loc(grand_col)
        status_cols = list(df_count.columns[1:grand_idx])
        df_data = df_count[df_count[row_col]!= 'Grand Total'].copy()
        df_data = df_data[df_data[row_col].notna() & (df_data[row_col]!='')]
        for c in status_cols + [grand_col]:
            df_data[c] = pd.to_numeric(df_data[c], errors='coerce').fillna(0)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("SYSTEMS", len(df_data))
        c2.metric("TOTAL WIR", int(df_data[grand_col].sum()))
        c3.metric("STATUS TYPES", len(status_cols))
        c4.metric("TOP SYSTEM", df_data.sort_values(grand_col, ascending=False).iloc[0][row_col] if len(df_data)>0 else "N/A")

        a,b = st.columns(2)
        with a:
            top10 = df_data.sort_values(grand_col, ascending=False).head(10)
            fig = px.bar(top10, x=grand_col, y=row_col, orientation='h', text=grand_col, title="Top 10 Systems")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
        with b:
            status_totals = [{'Status': sc, 'Count': df_data[sc].sum()} for sc in status_cols if df_data[sc].sum()>0]
            status_df = pd.DataFrame(status_totals)
            fig2 = px.pie(status_df, values='Count', names='Status', hole=0.5, title="STATUS WISE")
            fig2.update_traces(textinfo='percent+label')
            st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(df_data.head(30), use_container_width=True)
        final_df = df_data
    else:
        final_df = df
else:
    st.info("📌 Normal Excel File - Generic Analysis")
    # GENERIC LOGIC - Kisi bhi file ke liye
    text_cols = df.select_dtypes(include=['object']).columns.tolist()
    num_cols = df.select_dtypes(include='number').columns.tolist()

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("ROWS", len(df))
    c2.metric("COLS", len(df.columns))
    c3.metric("TEXT COLS", len(text_cols))
    c4.metric("NUMERIC", len(num_cols))

    # Auto best columns for chart
    chart_cols = text_cols[:4]
    if len(chart_cols) >= 2:
        a,b = st.columns(2)
        for i, col in enumerate(chart_cols[:2]):
            vc = df[col].astype(str).str.strip().replace('nan','').replace('None','')
            vc = vc[vc!=''].value_counts().head(8).reset_index()
            if len(vc)==0: continue
            vc.columns=[col,'Count']
            fig = px.pie(vc, values='Count', names=col, hole=0.5, title=f"{col}") if i==0 else px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} Top 8")
            if i==0: fig.update_traces(textinfo='percent+label')
            (a if i==0 else b).plotly_chart(fig, use_container_width=True)

    if len(chart_cols) >= 3:
        c,d = st.columns(2)
        for i, col in enumerate(chart_cols[2:4]):
            vc = df[col].astype(str).str.strip().replace('nan','').replace('None','')
            vc = vc[vc!=''].value_counts().head(10).reset_index()
            if len(vc)==0: continue
            vc.columns=[col,'Count']
            fig = px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} Top 10")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            (c if i==0 else d).plotly_chart(fig, use_container_width=True)

    st.dataframe(df.head(30), use_container_width=True)
    final_df = df

# --- FIXED DOWNLOAD FOR BOTH ---
def get_excel():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        final_df.to_excel(w, index=False, sheet_name='Data')
    out.seek(0)
    return out.getvalue()

st.divider()
st.download_button(f"📥 DOWNLOAD - {len(final_df)} Records", get_excel(), file_name=f"Universal_Report_{len(final_df)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
