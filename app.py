import streamlit as st
import pandas as pd
import io, os
import plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Dashboard - AUTO HEADER</h2>
</div>
""", unsafe_allow_html=True)

def find_best_header(file, sheet):
    best_row, max_cols = 0, 0
    for r in range(8):
        try:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet, header=r)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            cols = len([c for c in df.columns if 'Unnamed' not in str(c)])
            # Jahan Status, Row Labels, WIR jaisa mile wahi best hai
            col_str = ' '.join([str(c).lower() for c in df.columns])
            score = cols
            if any(x in col_str for x in ['status','row label','wir','system','originator','approved']):
                score += 10
            if score > max_cols:
                max_cols = score
                best_row = r
        except:
            pass
    return best_row

f = st.file_uploader("📂 Excel Upload Karo", type=["xlsx","xls"])

if not f:
    st.stop()

xls = pd.ExcelFile(f)
sheet = st.selectbox(f"Sheet Select Karo ({len(xls.sheet_names)} sheets)", xls.sheet_names)

# AUTO HEADER DETECT
auto_header = find_best_header(f, sheet)
st.info(f"🤖 Auto Detected Header Row: {auto_header} (Aap change kar sakte ho)")

header_row = st.number_input("Header Row Number", 0, 10, value=auto_header)

f.seek(0)
df = pd.read_excel(f, sheet_name=sheet, header=header_row)
df = df.dropna(how='all')
# Sirf puri khali columns hatao
df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False, regex=True) if 'Unnamed' in str(df.columns.tolist()) else [True]*len(df.columns)]

# Agar abhi bhi cols 0 hai to 0 se 7 tak try karo
if len(df.columns) == 0 or len(df.columns) < 2:
    for r in range(8):
        f.seek(0)
        temp = pd.read_excel(f, sheet_name=sheet, header=r)
        temp = temp.dropna(how='all').dropna(axis=1, how='all')
        if len(temp.columns) >= 3:
            df = temp
            header_row = r
            break

st.success(f"✅ Loaded: {len(df)} Records | {len(df.columns)} Columns | Header Row: {header_row}")
st.write("Columns:", list(df.columns)[:10])

# --- DASHBOARD ---
is_pivot = 'Row Labels' in df.columns

if is_pivot:
    # WIR PIVOT
    df_count = df[[c for c in df.columns if '.1' not in str(c)]]
    grand_col = [c for c in df_count.columns if 'Grand Total' in str(c)]
    grand_col = grand_col[0] if grand_col else None
    row_col = [c for c in df_count.columns if 'Row Labels' in str(c)]
    row_col = row_col[0] if row_col else df_count.columns[0]

    if grand_col:
        grand_idx = df_count.columns.get_loc(grand_col)
        status_cols = list(df_count.columns[1:grand_idx])
        df_data = df_count[(df_count[row_col]!='Grand Total') & (df_count[row_col].notna())].copy()
        for c in status_cols + [grand_col]:
            df_data[c] = pd.to_numeric(df_data[c], errors='coerce').fillna(0)

        c1,c2,c3 = st.columns(3)
        c1.metric("SYSTEMS", len(df_data))
        c2.metric("TOTAL WIR", int(df_data[grand_col].sum()))
        c3.metric("STATUS TYPES", len(status_cols))

        a,b = st.columns(2)
        with a:
            top10 = df_data.sort_values(grand_col, ascending=False).head(10)
            fig = px.bar(top10, x=grand_col, y=row_col, orientation='h', text=grand_col, title="Top 10 Systems - Status Wise")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
        with b:
            status_totals = [{'Status': sc, 'Count': int(df_data[sc].sum())} for sc in status_cols if df_data[sc].sum()>0]
            status_df = pd.DataFrame(status_totals)
            fig2 = px.pie(status_df, values='Count', names='Status', hole=0.5, title="STATUS WISE - A vs B vs C")
            fig2.update_traces(textinfo='percent+label')
            st.plotly_chart(fig2, use_container_width=True)
        final_df = df_data
    else:
        final_df = df
else:
    # NORMAL EXCEL - ANY FILE
    text_cols = df.select_dtypes(include=['object']).columns.tolist()
    # Agar object 0 hai to first columns ko text mano
    if len(text_cols)==0 and len(df.columns)>0:
        text_cols = list(df.columns[:4])

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("ROWS", len(df))
    c2.metric("COLS", len(df.columns))
    c3.metric("TEXT COLS", len(text_cols))
    c4.metric("NUMERIC", len(df.select_dtypes(include='number').columns))

    for i, col in enumerate(text_cols[:4]):
        try:
            vc = df[col].astype(str).str.strip()
            vc = vc[~vc.isin(['nan','None','', 'NaT'])].value_counts().head(10).reset_index()
            if len(vc)==0: continue
            vc.columns=[col,'Count']
            if i % 2 == 0:
                cols = st.columns(2)
            with cols[i%2]:
                if i==0:
                    fig = px.pie(vc, values='Count', names=col, hole=0.5, title=f"{col} - Breakdown")
                    fig.update_traces(textinfo='percent+label')
                else:
                    fig = px.bar(vc, x='Count', y=col, orientation='h', text='Count', title=f"{col} - Top 10")
                    fig.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
        except:
            pass
    final_df = df

st.markdown("### 📋 Data Preview")
st.dataframe(final_df.head(50), use_container_width=True)

def get_excel():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        final_df.to_excel(w, index=False, sheet_name='Data')
    out.seek(0)
    return out.getvalue()

st.download_button(f"📥 DOWNLOAD - {len(final_df)} Records", get_excel(), file_name=f"Report_{len(final_df)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
