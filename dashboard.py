
import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Fiber Prep Dashboard", layout="wide")
st.title("📊 Fiber Prep Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

def extract_drop_size(inventory):
    match = re.search(r"(\d{2,4})['’]\s?Drop", str(inventory))
    return match.group(1) + "'" if match else "Unknown"

if uploaded_file:
    try:
        # Load Excel file
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(xls, sheet_name=0)

        # Normalize columns
        df.columns = df.columns.str.strip()

        # Extract relevant fields
        df['Drop Size'] = df['INVENTORY ITEMS'].apply(extract_drop_size)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
        df['Tech'] = df['Tech'].astype(str).str.strip()

        # Sidebar filters
        st.sidebar.header("Filter Data")
        selected_date = st.sidebar.multiselect("Select Date", sorted(df['Date'].dropna().unique()))
        selected_tech = st.sidebar.multiselect("Select Tech", sorted(df['Tech'].dropna().unique()))
        selected_drop = st.sidebar.multiselect("Select Drop Size", sorted(df['Drop Size'].dropna().unique()))

        filtered_df = df.copy()
        if selected_date:
            filtered_df = filtered_df[filtered_df['Date'].isin(selected_date)]
        if selected_tech:
            filtered_df = filtered_df[filtered_df['Tech'].isin(selected_tech)]
        if selected_drop:
            filtered_df = filtered_df[filtered_df['Drop Size'].isin(selected_drop)]

        st.subheader("Filtered Results")
        st.dataframe(filtered_df, use_container_width=True)

        st.subheader("Summary")
        summary = filtered_df.groupby(['Date', 'Tech', 'Drop Size']).size().reset_index(name='Count')
        st.dataframe(summary)

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Upload an Excel file to begin.")
