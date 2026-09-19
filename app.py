import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Malik Pro Dashboard", layout="wide")

# --- PROFESSIONAL CSS ---
st.markdown("""
<style>
    .main { background-color: #f5f7fa; }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border-left: 5px solid #2F5597;
        text-align: center;
    }
    .metric-card h3 { color: #6c757d; font-size: 14px; margin: 0; }
    .metric-card h1 { color: #2F5597; font-size: 32px; font-weight: bold; margin: 5px 0; }
    .approved-card { border-left-color: #28a745; }
    .approved-card h1 { color: #28a745; }
    .review-card { border-left-color: #ffc107; }
    .review-card h1 { color: #d39e00; }
    .pending-card { border-left-color: #dc3545; }
    .pending-card h1 { color: #dc3545; }
    h1 { color: #1a2a4a; font-weight: 800; }
    .stDataFrame { background: white; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>📊 Malik Closeout - Professional Dashboard</h1><p style='color:gray'>Executive Summary Report</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📁 Koi bhi Excel / CSV upload karo - Auto Professional ban jayega", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    df_main = pd.DataFrame({
        'Category': ['Warranty Schedule', 'PDD', 'Spare Parts', 'O&M Manual', 'Training Schedule'],
        'QTY': [20, 14, 11, 15, 9],
        'Approved': [12, 8, 7, 10, 5],
        'Under Review': [3, 2, 1, 2, 2],
        'Pending': [5, 4, 3, 3, 2],
    })
else:
    df_main = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)

df_total = pd.DataFrame([['TOTAL', 69, 42, 10, 17]], columns=df_main.columns)
df_display = pd.concat([df_main, df_total], ignore_index=True)

# --- PROFESSIONAL METRIC CARDS ---
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f"<div class='metric-card'><h3>TOTAL QTY</h3><h1>69</h1><p style='color:gray;font-size:12px'>All Documents</p></div>", unsafe_allow_html=True)
with col2:
    st.markdown(f"<div class='metric-card approved-card'><h3>APPROVED</h3><h1>42</h1><p style='color:gray;font-size:12px'>✅ 61% Complete</p></div>", unsafe_allow_html=True)
with col3:
    st.markdown(f"<div class='metric-card review-card'><h3>UNDER REVIEW</h3><h1>10</h1><p style='color:gray;font-size:12px'>🕒 In Process</p></div>", unsafe_allow_html=True)
with col4:
    st.markdown(f"<div class='metric-card pending-card'><h3>PENDING</h3><h1>17</h1><p style='color:gray;font-size:12px'>⚠️ Action Required</p></div>", unsafe_allow_html=True)

st.write("")

# --- TABLE + CHART SIDE BY SIDE ---
c_left, c_right = st.columns([1.2, 1])

with c_left:
    st.subheader("📋 Detail Summary")
    # Styled dataframe
    def color_total(row):
        if row['Category'] == 'TOTAL':
            return ['background-color: #2F5597; color: white; font-weight: bold']*len(row)
        return ['']*len(row)
    st.dataframe(df_display.style.apply(color_total, axis=1), use_container_width=True, height=250)

with c_right:
    st.subheader("📈 Status Breakdown")
    fig = px.bar(df_main, x='Category', y=['Approved','Under Review','Pending'],
                 color_discrete_map={'Approved':'#28a745','Under Review':'#ffc107','Pending':'#dc3545'},
                 barmode='stack', template='plotly_white')
    fig.update_layout(height=280, margin=dict(l=10,r=10,t=10,b=80), legend=dict(orientation="h", y=-0.3), xaxis_tickangle=-20)
    st.plotly_chart(fig, use_container_width=True)

# --- DONUT CHART ---
st.subheader("🎯 Completion Overview")
col_d1, col_d2 = st.columns([1, 2])
with col_d1:
    donut_data = pd.DataFrame({'Status': ['Approved','Under Review','Pending'], 'Count': [42,10,17]})
    fig2 = px.pie(donut_data, values='Count', names='Status', hole=0.6,
                  color_discrete_map={'Approved':'#28a745','Under Review':'#ffc107','Pending':'#dc3545'})
    fig2.update_layout(height=250, showlegend=False, margin=dict(l=10,r=10,t=10,b=10))
    st.plotly_chart(fig2, use_container_width=True)
with col_d2:
    st.markdown("""
    <div style='background:white; padding:20px; border-radius:12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); height:250px'>
    <h4 style='color:#2F5597'>Insights</h4>
    <p>✅ <b>61% Approved</b> - Good progress</p>
    <p>🕒 <b>14% Under Review</b> - Awaiting feedback</p>
    <p>⚠️ <b>25% Pending</b> - Need immediate action on <b>Warranty Schedule (5)</b> and <b>PDD (4)</b></p>
    <p style='color:gray; font-size:12px; margin-top:15px'>Last Updated: Today | Project: Closeout</p>
    </div>
    """, unsafe_allow_html=True)

# --- PROFESSIONAL EXCEL ---
def make_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_display.to_excel(writer, index=False, sheet_name='SUMMARY', startrow=1)
        ws = writer.sheets['SUMMARY']
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        total_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        total_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        center = Alignment(horizontal='center', vertical='center')

        for col in range(1, len(df_display.columns)+1):
            c = ws.cell(row=2, column=col)
            c.fill = header_fill
            c.font = header_font
            c.border = border
            c.alignment = center

        for r in range(3, 3+len(df_display)):
            is_total = str(ws.cell(row=r, column=1).value).upper() == 'TOTAL'
            for col_idx in range(1, len(df_display.columns)+1):
                cell = ws.cell(row=r, column=col_idx)
                cell.border = border
                cell.alignment = center
                if is_total:
                    cell.fill = total_fill
                    cell.font = total_font
                else:
                    if col_idx == 3: cell.fill = green_fill
                    elif col_idx == 4: cell.fill = yellow_fill
                    elif col_idx == 5: cell.fill = red_fill

        for i in range(1, len(df_display.columns)+1):
            ws.column_dimensions[get_column_letter(i)].width = 18
    return output.getvalue()

st.divider()
st.download_button("📥 Professional CEO Report Download (Excel)", make_excel(), "Malik_CEO_Professional_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
