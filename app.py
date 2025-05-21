import streamlit as st
import pandas as pd

st.set_page_config(page_title="Payslip Sender", layout="centered")
st.title("📩 Payslip Generator and Email Sender")

# File uploader
uploaded_file = st.file_uploader("Upload Payroll Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        if 'Month' not in df.columns:
            st.error("Your Excel file must contain a 'Month' column.")
        else:
            available_months = df['Month'].dropna().unique()
            selected_month = st.selectbox("Select Month to Send Payslip", sorted(available_months))

            st.write("Sample Data for", selected_month)
            st.dataframe(df[df['Month'] == selected_month])
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
