import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Malik Ai Dashboard", page_icon="📊")
st.title("📊 Malik Ai Dashboard")
st.write("Excel Upload Karo - Auto Dashboard + Summary Pao")

file = st.file_uploader("Apni Excel / CSV File Upload Karo", type=["xlsx", "csv"])

if file:
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.success("File Upload Ho Gayi!")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Rows", len(df))
    col2.metric("Total Columns", len(df.columns))
    col3.metric("Numeric Columns", len(df.select_dtypes(include='number').columns))

    st.subheader("📈 Auto Dashboard")
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    if numeric_cols:
        fig = px.bar(df.head(20), x=df.columns[0], y=numeric_cols[0])
        st.plotly_chart(fig)

    st.subheader("📄 Data Table")
    st.dataframe(df.head(20))
