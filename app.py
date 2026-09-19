import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

st.set_page_config(page_title="Malik Closeout Dashboard", layout="wide")
st.title("📊 Malik Closeout Dashboard - Category Wise")

# --- CLOUSEOUT SUMMARY ---
st.subheader("CLOUSEOUT SUMMARY")

summary_data = {
    "Category": ["Warranty Schedule", "PDD- Photographic doc", "Spare Parts Schedule", "O&M Manual", "Training Schedule"],
    "QTY": [20, 14, 11, 15, 9],
    "Approved": [14, 0, 11, 12, 7],
    "Code D": [0, 14, 0, 0, 0],
    "Code C": [0, 0, 0, 0, 2],
    "Not Accepted": [0, 0, 0, 0, 0],
    "Under Review": [6, 0, 0, 3, 0],
    "Pending": [0, 0, 0, 0, 0],
    "Grand Total": [20, 14, 11, 15, 9]
}

df_sum = pd.DataFrame(summary_data)
total_row = pd.DataFrame([["TOTAL", 69, 44, 14, 2, 0, 9, 0, 69]], columns=df_sum.columns)
df_sum = pd.concat([df_sum, total_row], ignore_index=True)

st.dataframe(df_sum, use_container_width=True, hide_index=True)

# --- 5 CARDS ---
st.divider()
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Warranty", "20", "14 Approved")
col2.metric("PDD", "14", "14 Code D")
col3.metric("Spare Parts", "11", "11 Approved")
col4.metric("O&M Manual", "15", "12 Approved")
col5.metric("Training", "9", "7 Approved")

# --- CATEGORY WISE CHART ---
st.subheader("📊 Category Wise Chart Summary")
fig = go.Figure()
categories = summary_data["Category"]
fig.add_trace(go.Bar(name='Approved', x=categories, y=summary_data["Approved"], marker_color='#2ecc71'))
fig.add_trace(go.Bar(name='Code D', x=categories, y=summary_data["Code D"], marker_color='#e74c3c'))
fig.add_trace(go.Bar(name='Code C', x=categories, y=summary_data["Code C"], marker_color='#f39c12'))
fig.add_trace(go.Bar(name='Under Review', x=categories, y=summary_data["Under Review"], marker_color='#3498db'))
fig.update_layout(barmode='group', height=400)
st.plotly_chart(fig, use_container_width=True)

# --- EXCEL DOWNLOAD BUTTON ---
st.markdown("---")
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df_sum.to_excel(writer, index=False, sheet_name='CLOUSEOUT SUMMARY')

st.download_button(
    label="📥 CLOUSEOUT SUMMARY Excel me Download Karo",
    data=buffer.getvalue(),
    file_name="Malik_Closeout_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
