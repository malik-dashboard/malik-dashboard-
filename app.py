import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Malik Universal Dashboard", layout="wide")
st.title("📊 Malik Universal Dashboard - Koi bhi file")

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("Koi bhi Excel / CSV file upload karo", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    # Default data agar koi file nahi
    df = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule', 'TOTAL'],
        'QTY': [20, 14, 11, 15, 9, 69],
        'Approved': [15, 10, 8, 12, 7, 52],
        'Pending': [5, 4, 3, 3, 2, 17],
    })
    st.info("Default Closeout data dikh raha hai - Upar se apni file upload karo to ye badal jayega")
else:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    st.success(f"File upload ho gayi: {uploaded_file.name}")

# --- AUTO DASHBOARD ---
st.subheader("Data Preview")
st.dataframe(df, use_container_width=True)

# Metrics - auto
numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
if numeric_cols:
    cols = st.columns(len(numeric_cols[:4]))
    for i, col_name in enumerate(numeric_cols[:4]):
        cols[i].metric(col_name, int(df[col_name].sum()) if col_name!= 'TOTAL' else df[col_name].iloc[-1] if 'TOTAL' in df.astype(str).values else df[col_name].sum())

# Chart - auto
if len(numeric_cols) >= 1 and df.shape[0] > 1:
    st.subheader("Auto Chart")
    try:
        cat_col = df.select_dtypes(include=['object']).columns[0] if len(df.select_dtypes(include=['object']).columns) > 0 else df.columns[0]
        chart_df = df.set_index(cat_col)[numeric_cols[:2]] if len(numeric_cols)>=2 else df.set_index(cat_col)[numeric_cols[0]]
        st.bar_chart(chart_df)
    except:
        st.line_chart(df[numeric_cols])

# --- PROFESSIONAL EXCEL DOWNLOAD FUNCTION ---
def make_professional_excel(df_main):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_main.to_excel(writer, index=False, sheet_name='DATA', startrow=2)
        ws = writer.sheets['DATA']

        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        alt_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        thin_border = Border(left=Side(style='thin', color='2F5597'), right=Side(style='thin', color='2F5597'), top=Side(style='thin', color='2F5597'), bottom=Side(style='thin', color='2F5597'))
        center_align = Alignment(horizontal='center', vertical='center')

        last_col = chr(64 + len(df_main.columns))
        ws.merge_cells(f'A1:{last_col}1')
        ws['A1'] = f"PROFESSIONAL REPORT - {uploaded_file.name if uploaded_file else 'CLOSEOUT SUMMARY'}"
        ws['A1'].font = Font(bold=True, size=14, color="2F5597")
        ws['A1'].alignment = center_align

        for col in range(1, len(df_main.columns)+1):
            cell = ws.cell(row=3, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center_align

        for r in range(4, 4 + len(df_main)):
            for c in range(1, len(df_main.columns)+1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border
                cell.alignment = center_align
                if r % 2 == 0:
                    cell.fill = alt_fill

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20
    return output.getvalue()

# Download
st.divider()
excel_data = make_professional_excel(df)
st.download_button(
    label="📥 Professional Excel Download Karo - Jo bhi file ho",
    data=excel_data,
    file_name=f"Malik_Professional_{uploaded_file.name if uploaded_file else 'Report'}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
