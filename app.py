import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Malik Universal Dashboard", layout="wide")
st.title("📊 Malik Universal Dashboard - Koi bhi file")

uploaded_file = st.file_uploader("Koi bhi Excel / CSV file upload karo", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    df = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule'],
        'QTY': [20, 14, 11, 15, 9],
        'Approved': [15, 10, 8, 12, 7],
        'Pending': [5, 4, 3, 3, 2],
    })
    df_total = pd.DataFrame([['TOTAL', 69, 52, 17]], columns=df.columns)
    df_display = pd.concat([df, df_total], ignore_index=True)
    st.info("Default data - Upar se apni file upload karo")
else:
    df = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
    df_display = df
    st.success(f"File: {uploaded_file.name}")

st.subheader("Data Preview")
st.dataframe(df_display, use_container_width=True)

numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
if numeric_cols:
    c = st.columns(len(numeric_cols[:3]))
    for i, col in enumerate(numeric_cols[:3]):
        c[i].metric(col, int(df[col].sum()))

if len(df) > 0:
    st.subheader("Auto Chart")
    try:
        plot_df = df[df.iloc[:,0]!= 'TOTAL'] if 'TOTAL' in df.astype(str).values else df
        cat = plot_df.columns[0]
        st.bar_chart(plot_df.set_index(cat)[numeric_cols[:2]] if len(numeric_cols)>=2 else plot_df.set_index(cat)[numeric_cols[0]])
    except:
        st.bar_chart(df[numeric_cols])

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

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_main.columns))
        ws['A1'] = f"PROFESSIONAL REPORT - {uploaded_file.name if uploaded_file else 'CLOSEOUT'}"
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

        # Width fix - ye line theek hai ab
        for i in range(1, len(df_main.columns)+1):
            ws.column_dimensions[chr(64+i)].width = 22
    return output.getvalue()

st.divider()
excel_data = make_professional_excel(df_display)
st.download_button(
    label="📥 Professional Excel Download Karo",
    data=excel_data,
    file_name="Malik_Professional_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
