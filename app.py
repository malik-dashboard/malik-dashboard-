import streamlit as st
import pandas as pd
import io
import glob
import os
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Malik - Auto All Files", layout="wide")
st.title("📊 Malik - Auto Dashboard - All Files")
st.caption("Folder me jitni files hongi, sab ka dashboard khud banega")

# --- AUTO READ ALL FILES FROM REPO ---
# GitHub repo me jo bhi excel/csv files hongi, auto read
all_files = glob.glob("*.xlsx") + glob.glob("*.xls") + glob.glob("*.csv") + glob.glob("data/*.xlsx") + glob.glob("data/*.csv")

uploaded_file = st.file_uploader("📁 Ya Yahan Se Nayi File Upload Karo (Optional)", type=["xlsx","xls","csv"])

dfs = []

if uploaded_file:
    # Agar user ne upload ki hai to usko lo
    try:
        df_up = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, sheet_name=None)
        if isinstance(df_up, dict): df_up = list(df_up.values())[0]
        df_up['Source_File'] = uploaded_file.name
        dfs.append(df_up)
        st.success(f"Uploaded File Read: {uploaded_file.name}")
    except Exception as e:
        st.error(f"Upload error: {e}")

# Auto read all files from GitHub folder
if all_files:
    st.info(f"📂 Auto-Detected {len(all_files)} files in repo: {', '.join(all_files)}")
    for f in all_files:
        try:
            if f.endswith('.csv'):
                df_temp = pd.read_csv(f)
            else:
                df_temp = pd.read_excel(f, sheet_name=None)
                if isinstance(df_temp, dict):
                    df_temp = max(df_temp.values(), key=lambda x: x.shape[0])
            df_temp['Source_File'] = os.path.basename(f)
            dfs.append(df_temp)
        except:
            pass

# Agar koi file nahi mili to default Closeout
if not dfs:
    st.warning("Koi file nahi mili repo me - Default Closeout dikh raha hai. GitHub me files daalo to auto aayega")
    df_all = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule'],
        'QTY': [20, 14, 11, 15, 9],
        'Approved': [12, 8, 7, 10, 5],
        'Under Review': [3, 2, 1, 2, 2],
        'Pending': [5, 4, 3, 3, 2],
        'Source_File': ['Default']*5
    })
else:
    # Sab files ko ek saath jodo agar same columns hain to
    try:
        # Common columns wali files ko combine karo
        df_all = pd.concat(dfs, ignore_index=True, sort=False)
    except:
        df_all = dfs[0] # Agar combine nahi ho rahi to pehli dikhao

st.subheader(f"📋 All Files Data - Total {len(df_all)} Records")
st.dataframe(df_all, use_container_width=True, height=400)

# --- AUTO DASHBOARD - All Files Ka ---
num_cols = df_all.select_dtypes(include='number').columns.tolist()
obj_cols = df_all.select_dtypes(include='object').columns.tolist()
obj_cols = [c for c in obj_cols if c!= 'Source_File']

if num_cols:
    st.subheader("📊 All Files - Auto Summary")
    c = st.columns(min(len(num_cols), 4))
    for i, col_name in enumerate(num_cols[:4]):
        with c[i]:
            st.metric(col_name, f"{int(df_all[col_name].sum())}")

if obj_cols and num_cols:
    try:
        st.subheader("📈 All Files - Combined Chart")
        # Source_File se color karo taake pata chale kaunsi file se hai
        fig = px.bar(df_all.head(50), x=obj_cols[0], y=num_cols[0], color='Source_File' if 'Source_File' in df_all.columns else None, barmode='group', template='plotly_white')
        fig.update_layout(height=400, xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.write(e)

# File-wise count
if 'Source_File' in df_all.columns:
    st.subheader("📂 Files Breakdown")
    st.dataframe(df_all['Source_File'].value_counts(), use_container_width=True)

# --- EXCEL - All Files Ka ---
def make_all_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_all.to_excel(writer, index=False, sheet_name='ALL_FILES_REPORT', startrow=1)
        ws = writer.sheets['ALL_FILES_REPORT']
        hdr_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center = Alignment(horizontal='center', vertical='center')
        for col in range(1, len(df_all.columns)+1):
            cell = ws.cell(row=2, column=col)
            cell.fill = hdr_fill; cell.font = hdr_font; cell.border = border; cell.alignment = center
        for i in range(1, len(df_all.columns)+1):
            ws.column_dimensions[get_column_letter(i)].width = 18
    return output.getvalue()

st.divider()
st.download_button("📥 All Files Ka Professional Excel Download", make_all_excel(), "Malik_ALL_FILES_Auto_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.markdown("---")
st.caption("💡 Tip: GitHub repo me jaake 'Add file' -> 'Upload files' se jitni files daloge, ye dashboard khud sab ko auto read karke dashboard bana dega. Ek baar code lagao, baar baar file upload ka jhanjhat khatam!")
