import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Malik Universal Dashboard", layout="wide")
st.title("📊 Malik Universal Dashboard")

uploaded_file = st.file_uploader("Koi bhi Excel / CSV file upload karo", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    # SAHI DATA - Under Review ke saath
    df_main = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule'],
        'QTY': [20, 14, 11, 15, 9],
        'Approved': [12, 8, 7, 10, 5],
        'Under Review': [3, 2, 1, 2, 2],
        'Pending': [5, 4, 3, 3, 2],
    })
    # TOTAL sahi
    total_row = pd.DataFrame([['TOTAL', 69, 42, 10, 17]], columns=df_main.columns)
    df_display = pd.concat([df_main, total_row], ignore_index=True)
else:
    df_main = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    df_display = df_main
    # Agar file me Under Review nahi hai to bana do
    if 'Under Review' not in df_main.columns and 'Approved' in df_main.columns:
        df_main['Under Review'] = 0

st.subheader("Data Preview")
st.dataframe(df_display, use_container_width=True)

# Sahi Metrics - 4 columns
if uploaded_file is None:
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total QTY", "69")
    c2.metric("Approved", "42")
    c3.metric("Under Review", "10")
    c4.metric("Pending", "17")
else:
    st.write("Auto Metrics")
    nums = df_main.select_dtypes(include='number').columns.tolist()[:4]
    cols = st.columns(len(nums))
    for i, n in enumerate(nums):
        cols[i].metric(n, int(df_main[n].sum()))

# Sahi Chart - QTY nahi, sirf status
st.subheader("Chart - Status wise")
try:
    chart_df = df_main.set_index(df_main.columns[0])[['Approved','Under Review','Pending']]
    st.bar_chart(chart_df)
except Exception as e:
    st.bar_chart(df_main.select_dtypes(include='number'))

def make_professional_excel(df_to_save):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, index=False, sheet_name='SUMMARY', startrow=2)
        ws = writer.sheets['SUMMARY']

        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        total_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        review_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid") # Yellow for Under Review
        alt_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        total_font = Font(bold=True, color="FFFFFF", size=12)
        thin_border = Border(left=Side(style='thin', color='2F5597'), right=Side(style='thin', color='2F5597'), top=Side(style='thin', color='2F5597'), bottom=Side(style='thin', color='2F5597'))
        center = Alignment(horizontal='center', vertical='center')

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_to_save.columns))
        ws['A1'] = "PROFESSIONAL REPORT - CLOSEOUT WITH UNDER REVIEW"
        ws['A1'].font = Font(bold=True, size=14, color="2F5597")
        ws['A1'].alignment = center

        for col in range(1, len(df_to_save.columns)+1):
            cell = ws.cell(row=3, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center

        for r in range(4, 4 + len(df_to_save)):
            is_total = str(ws.cell(row=r, column=1).value).upper() == 'TOTAL'
            for c in range(1, len(df_to_save.columns)+1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border
                cell.alignment = center
                col_name = ws.cell(row=3, column=c).value
                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                else:
                    if col_name == 'Under Review':
                        cell.fill = review_fill
                    elif r % 2 == 0:
                        cell.fill = alt_fill

        for i in range(1, len(df_to_save.columns)+1):
            ws.column_dimensions[chr(64+i)].width = 18
    return output.getvalue()

st.divider()
excel_data = make_professional_excel(df_display)
st.download_button(
    label="📥 Sahi Professional Excel - Under Review ke saath",
    data=excel_data,
    file_name="Malik_Sahi_UnderReview_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
