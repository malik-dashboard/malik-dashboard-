import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Malik Closeout Dashboard", layout="wide")
st.title("📊 Malik Closeout Dashboard - Category Wise")

# --- CLOUSEOUT SUMMARY - Aapki photo jaisa ---
st.subheader("CLOUSEOUT SUMMARY")

summary_data = {
    "Category": ["Warranty Schedule", "PDO- Photographic doc", "Spare Parts Schedule", "O&M Manual", "Training Schedule"],
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
# Total row
total_row = pd.DataFrame([["TOTAL", 69, 44, 14, 2, 0, 9, 0, 69]], columns=df_sum.columns)
df_sum = pd.concat([df_sum, total_row])

st.dataframe(df_sum, use_container_width=True, hide_index=True)

# --- 5 CARDS ---
st.divider()
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Warranty", "20", "14 Approved")
col2.metric("PDO", "14", "14 Code D")
col3.metric("Spare Parts", "11", "11 Approved")
col4.metric("O&M Manual", "15", "12 Approved")
col5.metric("Training", "9", "7 Approved")

# --- CATEGORY WISE CHART ---
st.subheader("📈 Category Wise Chart Summary")
fig = go.Figure()
fig.add_trace(go.Bar(x=df_sum["Category"][:-1], y=df_sum["Approved"][:-1], name="Approved", marker_color="#2ecc71"))
fig.add_trace(go.Bar(x=df_sum["Category"][:-1], y=df_sum["Code D"][:-1], name="Code D", marker_color="#e74c3c"))
fig.add_trace(go.Bar(x=df_sum["Category"][:-1], y=df_sum["Code C"][:-1], name="Code C", marker_color="#f39c12"))
fig.add_trace(go.Bar(x=df_sum["Category"][:-1], y=df_sum["Under Review"][:-1], name="Under Review", marker_color="#3498db"))
fig.update_layout(barmode='group', height=500, xaxis_tickangle=-20)
st.plotly_chart(fig, use_container_width=True)

# --- EXCEL UPLOAD ---
st.divider()
st.subheader("📄 Original Excel Upload")
file = st.file_uploader("Apni 107 rows wali Excel yahan upload karo", type=["xlsx","csv"])
if file:
    df = pd.read_excel(file) if file.name.endswith("xlsx") else pd.read_csv(file)
    st.success("File Upload Ho Gayi!")
    st.dataframe(df, use_container_width=True, height=500)
