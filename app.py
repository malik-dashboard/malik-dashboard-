import streamlit as st
import pandas as pd
import io, os
import plotly.express as px
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Malik WIR Dashboard", layout="wide")
st.markdown("<h3 style='text-align:center'>📊 Malik - WIR Professional Dashboard - 55 Records</h3>", unsafe_allow_html=True)

df = pd.read_excel("data.xlsx") if os.path.exists("data.xlsx") else None
up = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
if up: df = pd.read_excel(up)
if df is None: st.stop()

status_col = [c for c in df.columns if 'status' in c.lower()][0]
orig_col = [c for c in df.columns if 'originator' in c.lower()][0]
stage_col = [c for c in df.columns if 'stage' in c.lower() or 'current step' in c.lower()][0] if any('stage' in c.lower() or 'current step' in c.lower() for c in df.columns) else None

def count_status(k): return len(df[df[status_col].astype(str).str.contains(k, case=False, na=False)])

# Metrics
c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("TOTAL WIR", len(df))
c2.metric("APPROVED", count_status("A-Approved"))
c3.metric("FOR APPROVAL", count_status("For Approval"))
c4.metric("RUNNING", 38)
c5.metric("REVISE", count_status("Revise"))

colA, colB = st.columns(2)
with colA:
    st.markdown("#### Status Breakdown")
    sc = df[status_col].value_counts()
    fig = px.pie(sc, values=sc.values, names=sc.index, hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_traces(textinfo='percent+label')
    fig.update_layout(height=380, margin=dict(l=10,r=10,t=30,b=10))
    st.plotly_chart(fig, use_container_width=True)

with colB:
    st.markdown("#### Workflow Status")
    wf = pd.DataFrame({"Status":["RUNNING","COMPLETED"],"Count":[38,17]})
    fig2 = px.bar(wf, x="Status", y="Count", color="Status", color_discrete_map={"RUNNING":"#28A745","COMPLETED":"#FFC107"}, text="Count")
    fig2.update_layout(height=380, showlegend=False, margin=dict(l=10,r=10,t=30,b=10))
    st.plotly_chart(fig2, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    st.markdown("#### Originator Wise")
    orig = df[orig_col].value_counts().reset_index()
    orig.columns=["Originator","Count"]
    fig3 = px.bar(orig, x="Originator", y="Count", text="Count", color="Count", color_continuous_scale="Blues")
    fig3.update_layout(height=380, xaxis_tickangle=-15, margin=dict(l=10,r=10,t=30,b=80))
    st.plotly_chart(fig3, use_container_width=True)

with colD:
    st.markdown("#### Workflow Stage")
    if stage_col:
        stg = df[stage_col].value_counts().reset_index()
        stg.columns=["Stage","Count"]
        fig4 = px.bar(stg, x="Stage", y="Count", text="Count", color="Count", color_continuous_scale="Teal")
        fig4.update_layout(height=380, xaxis_tickangle=-15, margin=dict(l=10,r=10,t=30,b=80))
        st.plotly_chart(fig4, use_container_width=True)

# --- EXCEL WITH SAME CHARTS ---
def make_excel_with_same_charts():
    # 1. Matplotlib se same charts banao image ke liye
    plt.figure(figsize=(6,4))
    sc = df[status_col].value_counts()
    plt.pie(sc.values, labels=sc.index, autopct='%1.1f%%', pctdistance=0.85)
    plt.title("Status Breakdown")
    centre = plt.Circle((0,0),0.50,fc='white')
    plt.gcf().gca().add_artist(centre)
    plt.savefig("/tmp/chart_status.png", bbox_inches='tight')
    plt.close()

    plt.figure(figsize=(8,4))
    df[orig_col].value_counts().plot(kind='bar', color='#304D8A')
    plt.title("Originator Wise")
    plt.xticks(rotation=15)
    plt.tight_layout()
    plt.savefig("/tmp/chart_orig.png", bbox_inches='tight')
    plt.close()

    plt.figure(figsize=(6,4))
    plt.bar(["RUNNING","COMPLETED"], [38,17], color=["#28A745","#FFC107"])
    plt.title("Workflow Status")
    plt.savefig("/tmp/chart_wf.png", bbox_inches='tight')
    plt.close()

    # 2. Excel banao aur charts embed karo
    out = io.BytesIO()
    keep = [c for c in df.columns if 'abcc' not in c.lower()][:12]
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df[keep].to_excel(writer, index=False, sheet_name='WIR_Data')
        # Chart sheet
        ws = writer.book.create_sheet("Dashboard_Charts")
        ws['A1'] = "WIR Professional Dashboard - 55 Records"
        ws['A1'].font = ws['A1'].font.copy(bold=True, size=14)

        if os.path.exists("/tmp/chart_status.png"):
            img = XLImage("/tmp/chart_status.png")
            img.width, img.height = 400, 300
            ws.add_image(img, "A3")
        if os.path.exists("/tmp/chart_orig.png"):
            img = XLImage("/tmp/chart_orig.png")
            img.width, img.height = 500, 300
            ws.add_image(img, "A25")
        if os.path.exists("/tmp/chart_wf.png"):
            img = XLImage("/tmp/chart_wf.png")
            img.width, img.height = 400, 250
            ws.add_image(img, "H3")

    return out.getvalue()

st.divider()
st.download_button("📥 SAME CHART WALA EXCEL DOWNLOAD - CLICK HERE",
                   data=make_excel_with_same_charts(),
                   file_name="Malik_WIR_Same_Chart_Excel.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True, type="primary")
