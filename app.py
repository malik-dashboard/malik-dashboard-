# --- PROFESSIONAL EXCEL DOWNLOAD WITH BORDERS & COLOURS ---
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_sum.to_excel(writer, index=False, sheet_name='CLOSEOUT SUMMARY', startrow=2)
    ws = writer.sheets['CLOSEOUT SUMMARY']
    
    # --- Styles ---
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
    
    # Title
    ws.merge_cells('A1:D1')
    ws['A1'] = "CLOSEOUT SUMMARY REPORT - PROFESSIONAL"
    ws['A1'].font = Font(bold=True, size=14, color="2F5597")
    ws['A1'].alignment = center_align
    
    # Header Row (Row 3)
    for col in range(1, 5):
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = center_align
    
    # Data Rows
    for r_idx in range(4, 4 + len(df_sum)):
        for c_idx in range(1, 5):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.alignment = center_align
            cell.font = total_font if r_idx == 4 + len(df_sum) -1 else normal_font
            
            if r_idx == 4 + len(df_sum) -1: # TOTAL row
                cell.fill = total_fill
            elif r_idx % 2 == 0:
                cell.fill = alt_fill
            else:
                cell.fill = white_fill
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15

st.download_button(
    label="📥 PROFESSIONAL Excel Download Karo",
    data=output.getvalue(),
    file_name="Malik_Professional_Closeout_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
