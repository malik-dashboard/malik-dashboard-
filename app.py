import streamlit as st
import pandas as pd
import io, os
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Malik WIR Dashboard", layout="wide", page_icon="📊")
st.markdown("<h2 style='text-align:center'>📊 Malik - WIR Professional Dashboard - 55 Records</h2>", unsafe_allow_html=True)

df = None
if os.path.exists("data.xlsx"):
    df = pd.read_excel("data.xlsx")
up = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
if up:
    df = pd.read_excel(up)
if df is None:
    st.stop()

# Clean columns
status_col = [c for c in df.columns if 'status' in c.lower()][0]
orig_col = [c for c in df.columns if 'originator' in c.lower()][0]

# Top Metrics
c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("TOTAL WIR", len(df), delta="All Records")
c2.metric("APPROVED", len(df[df[status_col].astype(str).str.contains('A-Approved', case=False, na=False)]))
c3.metric("FOR APPROVAL", len(df[df[status_col].astype(str).str.contains('For Approval', case=False, na=False)]))
c4.metric("RUNNING", 38)
c5.metric("REVISE", 5)

st.divider()

# PROFESSIONAL CHARTS
colA, colB = st.columns(2)

with colA:
    st.markdown("#### Status Breakdown")
    status_counts = df[status_col].value_counts()
    fig = px.pie(status_counts, values=status_counts.values, names=status_counts.index, hole=0.5,
                 color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_layout(height=400, showlegend=True, legend=dict(orientation="v"))
    fig.update_traces(textinfo='percent+label')
    st.plotly_chart(fig, use_container_width=True)

with colB:
    st.markdown("#### Workflow Status")
    wf = pd.DataFrame({"Status":["RUNNING","COMPLETED"],"Count":[38,17]})
    fig2 = px.bar(wf, x="Status", y="Count", color="Status",
                  color_discrete_map={"RUNNING":"#28A745","COMPLETED":"#FFC107"}, text="Count")
    fig2.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    st.markdown("#### Originator Wise")
    orig = df[orig_col].value_counts().reset_index()
    orig.columns=["Originator","Count"]
    fig3 = px.bar(orig, x="Originator", y="Count", color="Count",
                  color_continuous_scale="Blues", text="Count")
    fig3.update_layout(height=400, xaxis_tickangle=-20)
    st.plotly_chart(fig3, use_container_width=True)

with colD:
    st.markdown("#### Workflow Stage")
    stage_col = [c for c in df.columns if 'workflow stage' in c.lower() or 'current step' in c.lower()]
    if stage_col:
        stg = df[stage_col[0]].value_counts().reset_index()
        stg.columns=["Stage","Count"]
        fig4 = px.bar(stg, x="Stage", y="Count", color="Count", color_continuous_scale="Teal", text="Count")
        fig4.update_layout(height=400, xaxis_tickangle=-20)
        st.plotly_chart(fig4, use_container_width=True)

# EXCEL DOWNLOAD - Clean
def make_excel():
    out = io.BytesIO()
    keep = [c for c in df.columns if 'abcc' not in c.lower()][:15]
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df[keep].to_excel(writer, index=False, sheet_name='WIR_DATA')
        df[status_col].value_counts().reset_index().to_excel(writer, index=False, sheet_name='STATUS_CHART')
        df[orig_col].value_counts().reset_index().to_excel(writer, index=False, sheet_name='ORIGINATOR_CHART')
    return out.getvalue()

st.divider()
st.download_button("📥 PROFESSIONAL EXCEL REPORT DOWNLOAD (Chart Data Ke Saath)",
                   data=make_excel(),
                   file_name="Malik_Professional_WIR_Report.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True, type="primary")
