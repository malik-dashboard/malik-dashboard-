import streamlit as st
import pandas as pd
import io, plotly.express as px

st.set_page_config(page_title="Universal Dashboard", layout="wide", page_icon="📊")

st.markdown("""
<div style='text-align:center; background: linear-gradient(90deg, #0A1931 0%, #185ADB 100%); padding:15px; border-radius:10px; color:white'>
<h2 style='margin:0; color:white'>📊 Universal Data Analytics Dashboard</h2>
<p style='margin:0'>Pivot & Raw Data - Both Supported</p>
</div>
""", unsafe_allow_html=True)

f = st.file_uploader("📂 Excel File Upload Karo", type=["xlsx","xls"])

if not f:
    st.stop()

# --- ALL SHEETS DIKHAO ---
xls = pd.ExcelFile(f)
st.write(f"Is file me {len(xls.sheet_names)} Sheets hain:", xls.sheet_names)

sheet = st.selectbox("Kaunsi Sheet ka Dashboard banana hai? (WIR Log wali sheet select karo)", xls.sheet_names)

raw = pd.read_excel(f, sheet_name=sheet, header=None, nrows=10)
st.dataframe(raw, use_container_width=True)

header_row = st.number_input("Header Row Number?", 0, 9, 5 if 'Row Labels' in str(raw.values) else 0)

df = pd.read_excel(f, sheet_name=sheet, header=header_row)
df = df.dropna(how='all').dropna(axis=1, how='all')
df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False, regex=True)]

st.success(f"Loaded: {len(df)} Records | Columns: {list(df.columns)[:6]}")

# Check if Pivot file hai
if 'Row Labels' in df.columns or 'Count of STATUS' in str(df.columns):
    st.warning("⚠️ Ye PIVOT SUMMARY file hai - Raw WIR Log nahi! Status upar heading me hai - Niche se dekh lo:")
    # Pivot ka chart
    if 'Row Labels' in df.columns:
        first_col = 'Row Labels'
        # A-Approved, B-Approved wale columns hi Status hain
        status_cols = [c for c in df.columns if 'Approved' in str(c) or 'Revise' in str(c) or 'Accepted' in str(c)]
        st.write("Is Pivot me ye Status mile:", status_cols)
        
        # Bar chart from pivot
        plot_df = df[df[first_col].notna() & (df[first_col]!='Grand Total')].head(20)
        if not plot_df.empty and status_cols:
            melted = plot_df.melt(id_vars=[first_col], value_vars=status_cols, var_name='Status', value_name='Count')
            melted = melted[melted['Count']>0]
            fig = px.bar(melted, x=first_col, y='Count', color='Status', title="Status by System / Level - Pivot Data", barmode='stack')
            st.plotly_chart(fig, use_container_width=True)

st.dataframe(df.head(30), use_container_width=True)

out = io.BytesIO()
with pd.ExcelWriter(out, engine='openpyxl') as w:
    df.to_excel(w, index=False, sheet_name='Data')
st.download_button("📥 DOWNLOAD", out.getvalue(), file_name="Report.xlsx", use_container_width=True, type="primary")
