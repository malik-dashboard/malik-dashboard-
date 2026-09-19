import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Malik Universal Dashboard", layout="wide")
st.title("📊 Malik Universal Dashboard")

uploaded_file = st.file_uploader("Koi bhi Excel / CSV file upload karo", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    df_main = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule'],
        'QTY': [20, 14, 11, 15, 9],
        'Approved': [12, 8, 7, 10, 5],
        'Under Review': [3, 2, 1, 2, 2],
        'Pending': [5, 4, 3, 3, 2],
    })
    total_row = pd.DataFrame([['TOTAL', 69, 42, 10, 17]], columns=df_main.columns)
    df_display = pd.concat([df_main, total_row], ignore_index=True)
else:
    df_main = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    df_display = df_main

st.dataframe(df_display, use_container_width=True)

c1,c2,c3,c4 = st.columns(4)
if uploaded_file is None:
    c1.metric("QTY", 69)
    c2.metric("Approved", 42)
    c3.metric("Under Review", 10)
    c4.metric("Pending", 17)

st.bar_chart(df_main.set_index(df_main.columns[0])[['Approved','Under Review','Pending']])

# --- 100% WORKING EXCEL ---
def make_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_display.to_excel(writer, index=False, sheet_name='SUMMARY', startrow=1)
        ws = writer.sheets['SUMMARY']

        # Style objects - safe colors
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        total_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        alt_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        header_font = Font(bold=True, color="FFFFFF", size=11)
        total_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center = Alignment(horizontal='center', vertical='center')

        # Header row is now row 2 (because startrow=1)
        for col_num in range(1, len(df_display.columns) + 1):
            cell = ws.cell(row=2, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center

        # Data rows start at row 3
        for r_idx in range(3, 3 + len(df_display)):
            category = str(ws.cell(row=r_idx, column=1).value).upper()
            is_total = category == 'TOTAL'
            for c_idx in range(1, len(df_display.columns) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = border
                cell.alignment = center
                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                else:
                    # Under Review column yellow
                    if c_idx == 4: # 4th column is Under Review
                        cell.fill = yellow_fill
                    elif r_idx % 2 == 0:
                        cell.fill = alt_fill

        # Column width - FIXED METHOD
        for col_idx in range(1, len(df_display.columns) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 18

    return output.getvalue()

st.divider()
excel_file = make_excel()
st.download_button(
    label="📥 FIXED Excel Download Karo - Ab Khulega",
    data=excel_file,
    file_name="Malik_Fixed_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
