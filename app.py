import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Malik Universal Dashboard", layout="wide")
st.title("📊 Malik Universal Dashboard")

uploaded_file = st.file_uploader("Koi bhi Excel / CSV file upload karo", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    # Sahi data - TOTAL alag
    df_main = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule'],
        'QTY': [20, 14, 11, 15, 9],
        'Approved': [15, 10, 8, 12, 7],
        'Pending': [5, 4, 3, 3, 2],
    })
    # TOTAL calculate karke alag se
    total_row = pd.DataFrame([['TOTAL', df_main['QTY'].sum(), df_main['Approved'].sum(), df_main['Pending'].sum()]], columns=df_main.columns)
    df_display = pd.concat([df_main, total_row], ignore_index=True)
else:
    df_main = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    df_display = df_main

st.subheader("Data Preview")
st.dataframe(df_display, use_container_width=True)

# Sahi Metrics - TOTAL ko minus karke
if uploaded_file is None:
    c1,c2,c3 = st.columns(3)
    c1.metric("QTY", int(df_main['QTY'].sum()))
    c2.metric("Approved", int(df_main['Approved'].sum()))
    c3.metric("Pending", int(df_main['Pending'].sum()))
else:
    nums = df_main.select_dtypes(include='number').columns.tolist()[:3]
    if nums:
        cols = st.columns(len(nums))
        for i, n in enumerate(nums):
            cols[i].metric(n, int(df_main[n].sum()))

# Chart - TOTAL ke baghair
st.subheader("Chart")
try:
    chart_data = df_main # TOTAL nahi hai isme
    st.bar_chart(chart_data.set_index(chart_data.columns[0]).select_dtypes(include='number'))
except:
    pass

def make_professional_excel(df_to_save, df_for_total_check):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, index=False, sheet_name='SUMMARY', startrow=2)
        ws = writer.sheets['SUMMARY']

        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        total_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        alt_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        total_font = Font(bold=True, color="FFFFFF", size=12)
        normal_font = Font(size=11)
        thin_border = Border(left=Side(style='thin', color='2F5597'), right=Side(style='thin', color='2F5597'), top=Side(style='thin', color='2F5597'), bottom=Side(style='thin', color='2F5597'))
        center = Alignment(horizontal='center', vertical='center')

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_to_save.columns))
        ws['A1'] = "PROFESSIONAL REPORT - CLOSEOUT"
        ws['A1'].font = Font(bold=True, size=14, color="2F5597")
        ws['A1'].alignment = center

        # Header
        for col in range(1, len(df_to_save.columns)+1):
            cell = ws.cell(row=3, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center

        # Data with TOTAL blue
        for r in range(4, 4 + len(df_to_save)):
            is_total = str(ws.cell(row=r, column=1).value).upper() == 'TOTAL'
            for c in range(1, len(df_to_save.columns)+1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border
                cell.alignment = center
                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                else:
                    cell.font = normal_font
                    if r % 2 == 0:
                        cell.fill = alt_fill

        for i in range(1, len(df_to_save.columns)+1):
            ws.column_dimensions[chr(64+i)].width = 20
    return output.getvalue()

st.divider()
excel_data = make_professional_excel(df_display, df_main)
st.download_button(
    label="📥 Sahi Professional Excel Download Karo",
    data=excel_data,
    file_name="Malik_Sahi_Professional_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
