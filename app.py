import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Malik - WIR Dashboard", layout="wide")

# --- PROFESSIONAL CSS ---
st.markdown("""
<style>
.metric-card { background:white; padding:20px; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.08); border-left:5px solid #2F5597; text-align:center; }
.metric-card h3 { color:#6c757d; font-size:12px; margin:0; }
.metric-card h1 { color:#2F5597; font-size:28px; font-weight:bold; margin:5px 0; }
.green { border-left-color:#28a745; } .green h1 { color:#28a745; }
.yellow { border-left-color:#ffc107; } .yellow h1 { color:#d39e00; }
.red { border-left-color:#dc3545; } .red h1 { color:#dc3545; }
.blue { border-left-color:#17a2b8; } .blue h1 { color:#17a2b8; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Malik - WIR Inspection Dashboard")
st.caption("Form L_106 - Professional Status Report")

uploaded_file = st.file_uploader("📁 File Upload Karo", type=["xlsx","xls","csv"])

if uploaded_file is None:
    st.info("File upload karo - Form L wali")
    st.stop()

try:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    df = df.dropna(how='all')
    # Clean column names
    df.columns = df.columns.str.strip()
except Exception as e:
    st.error(f"Error: {e}")
    st.stop()

st.success(f"✅ Read: {uploaded_file.name} | Total Records: {len(df)}")

# --- AUTO DETECT COLUMNS from your screenshot ---
status_col = next((c for c in df.columns if 'Status' == c or c.lower()=='status'), None)
wf_status_col = next((c for c in df.columns if 'Workflow Status' in c), None)
wf_stage_col = next((c for c in df.columns if 'Workflow Stage' in c), None)
originator_col = next((c for c in df.columns if 'Originator' in c and 'Org' not in c), None)

# --- METRICS ---
total = len(df)
approved = len(df[df[status_col].astype(str).str.contains('Approved', case=False, na=False)]) if status_col else 0
for_auth = len(df[df[status_col].astype(str).str.contains('Authorisation', case=False, na=False)]) if status_col else 0
running = len(df[df[wf_status_col].astype(str).str.contains('RUNNING', case=False, na=False)]) if wf_status_col else 0
completed = len(df[df[wf_status_col].astype(str).str.contains('COMPLETED', case=False, na=False)]) if wf_status_col else 0
revise = len(df[df[status_col].astype(str).str.contains('Revise', case=False, na=False)]) if status_col else 0

c1,c2,c3,c4,c5 = st.columns(5)
c1.markdown(f"<div class='metric-card'><h3>TOTAL WIR</h3><h1>{total}</h1></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='metric-card green'><h3>APPROVED</h3><h1>{approved}</h1></div>", unsafe_allow_html=True)
c3.markdown(f"<div class='metric-card yellow'><h3>FOR AUTHORISATION</h3><h1>{for_auth}</h1></div>", unsafe_allow_html=True)
c4.markdown(f"<div class='metric-card blue'><h3>RUNNING</h3><h1>{running}</h1></div>", unsafe_allow_html=True)
c5.markdown(f"<div class='metric-card red'><h3>REVISE & SUBMIT</h3><h1>{revise}</h1></div>", unsafe_allow_html=True)

st.write("")

# --- CHARTS ---
left, right = st.columns(2)

with left:
    if status_col:
        st.subheader("📈 Status Breakdown")
        status_counts = df[status_col].value_counts().reset_index()
        status_counts.columns = ['Status','Count']
        fig1 = px.pie(status_counts, values='Count', names='Status', hole=0.5, template='plotly_white', color_discrete_sequence=px.colors.qualitative.Set2)
        fig1.update_layout(height=350)
        st.plotly_chart(fig1, use_container_width=True)

with right:
    if wf_status_col:
        st.subheader("📊 Workflow Status")
        wf_counts = df[wf_status_col].value_counts().reset_index()
        wf_counts.columns = ['Workflow','Count']
        fig2 = px.bar(wf_counts, x='Workflow', y='Count', color='Workflow', template='plotly_white', color_discrete_sequence=['#28a745','#ffc107','#dc3545'])
        fig2.update_layout(height=350, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

col3, col4 = st.columns(2)
with col3:
    if originator_col:
        st.subheader("👷 Originator Wise")
        orig_counts = df[originator_col].value_counts().head(10).reset_index()
        orig_counts.columns = ['Originator','Count']
        fig3 = px.bar(orig_counts, x='Originator', y='Count', template='plotly_white', color_discrete_sequence=['#2F5597'])
        st.plotly_chart(fig3, use_container_width=True)

with col4:
    if wf_stage_col:
        st.subheader("🏗️ Workflow Stage")
        stage_counts = df[wf_stage_col].value_counts().reset_index()
        stage_counts.columns = ['Stage','Count']
        fig4 = px.bar(stage_counts, x='Stage', y='Count', template='plotly_white', color_discrete_sequence=['#17a2b8'])
        st.plotly_chart(fig4, use_container_width=True)

# --- TABLE WITH FILTER ---
st.subheader("📋 Detail Data - 55 Records")
# Filter
if status_col:
    filter_status = st.multiselect(f"Filter by {status_col}", options=df[status_col].dropna().unique(), default=df[status_col].dropna().unique())
    df_filtered = df[df[status_col].isin(filter_status)]
else:
    df_filtered = df

st.dataframe(df_filtered, use_container_width=True, height=400)

# --- PROFESSIONAL EXCEL - 100% WORKING ---
def make_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtered.to_excel(writer, index=False, sheet_name='WIR_REPORT')
    return output.getvalue()

st.divider()
excel_data = make_excel()
st.download_button(
    label="📥 Professional WIR Report Download - Click Here",
    data=excel_data,
    file_name="Malik_WIR_Professional_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary"
)
