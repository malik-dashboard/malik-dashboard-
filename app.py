import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Malik Closeout Dashboard", layout="wide")
st.title("📊 Malik Closeout Dashboard")

# --- DATA ---
data = {
    'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule', 'TOTAL'],
    'QTY': [20, 14, 11, 15, 9, 69],
    'Approved': [15, 10, 8, 12, 7, 52],
    'Pending': [5, 4, 3, 3, 2, 17],
}
df_sum = pd.DataFrame(data)

# --- DASHBOARD ---
col1, col2, col3 = st.columns(3)
col1.metric("Total QTY", "69")
col2.metric("Approved", "52")
col3.metric("Pending", "17")

st.subheader("Closeout Summary")
st.dataframe(df_sum, use_container_width=True)
st.bar_chart(df_sum[df_sum['Category'] != 'TOTAL'].set_index('Category')[['Approved', 'Pending']])

# --- PROFESSIONAL EXCEL DOWNLOAD ---
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_sum.to_excel(writer, index=False, sheet_name='CLOSEOUT SUMMARY', startrow=2)
    ws = writer.sheets['CLOSEOUT SUMMARY']
    
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    total_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    alt_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_font = Font(bold=True, color="FFFFFF", size=11)
    normal_font = Font(size=11)
    
    thin_border = Border(
        left=Side(style='thin', color='2F5597'),
        right=Side(style='thin', color='2F5597'),
        top=Side(style='thin', color='2F5597'),
        bottom=Side(style='thin', color='2F5597')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells('A1:D1')
    ws['A1'] = "CLOSEOUT SUMMARY REPORT - PROFESSIONAL"
    ws['A1'].font = Font(bold=True, size=14, color="2F5597")
    ws['A1'].alignment = center_align
    
    for col in range(1, 5):
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_align
    
    for r_idx in range(4, 4 + len(df_sum)):
        for c_idx in range(1, 5):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.alignment = center_align
            cell.font = total_font if r_idx == 4 + len(df_sum) -1 else normal_font
            if r_idx == 4 + len(df_sum) -1:
                cell.fill = total_fill
            elif r_idx % 2 == 0:
                cell.fill = alt_fill
            else:
                cell.fill = white_fill
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15

st.divider()
st.download_button(
    label="📥 PROFESSIONAL Excel Download Karo",
    data=output.getvalue(),
    file_name="Malik_Professional_Closeout_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
