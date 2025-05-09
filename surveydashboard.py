
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
        # Load and clean the Excel data
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df.columns = df.columns.str.strip()  # clean column headers

        # Ensure necessary columns exist
        required_columns = ['Date', 'Tech', 'INVENTORY ITEMS']
        if not all(col in df.columns for col in required_columns):
            st.error("Missing required columns in the Excel file: 'Date', 'Tech', or 'INVENTORY ITEMS'")
        else:
            # Extract and clean data
            df['Drop Size'] = df['INVENTORY ITEMS'].apply(extract_drop_size)
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
            df['Tech'] = df['Tech'].astype(str).str.strip()

            # Sidebar filters
            st.sidebar.header("Filter Data")
            selected_dates = st.sidebar.multiselect("Select Date(s)", sorted(df['Date'].dropna().unique()))
            selected_techs = st.sidebar.multiselect("Select Tech(s)", sorted(df['Tech'].dropna().unique()))
            selected_drops = st.sidebar.multiselect("Select Drop Size(s)", sorted(df['Drop Size'].dropna().unique()))

            filtered_df = df.copy()
            if selected_dates:
                filtered_df = filtered_df[filtered_df['Date'].isin(selected_dates)]
            if selected_techs:
                filtered_df = filtered_df[filtered_df['Tech'].isin(selected_techs)]
            if selected_drops:
                filtered_df = filtered_df[filtered_df['Drop Size'].isin(selected_drops)]

            st.subheader("Filtered Results")
            st.dataframe(filtered_df, use_container_width=True)

            st.subheader("Summary by Date, Tech, and Drop Size")
            summary = filtered_df.groupby(['Date', 'Tech', 'Drop Size']).size().reset_index(name='Count')
            st.dataframe(summary)

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Upload an Excel file to begin.")
