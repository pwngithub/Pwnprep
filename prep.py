
import streamlit as st
import pandas as pd
import re
import altair as alt

st.set_page_config(page_title="Fiber Prep Dashboard", layout="wide")
st.title("📊 Fiber Prep Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

def extract_drop_size(inventory):
    match = re.search(r"(\d{2,4})['’]\s?Drop", str(inventory))
    return match.group(1) + "'" if match else "Unknown"

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df.columns = df.columns.str.strip()

        required_columns = ['Date', 'Tech', 'INVENTORY ITEMS']
        if not all(col in df.columns for col in required_columns):
            st.error("Missing required columns: 'Date', 'Tech', or 'INVENTORY ITEMS'")
        else:
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

            st.markdown("---")
            st.header("📊 Visualizations")

            # 1. Installs over time
            installs_over_time = filtered_df.groupby('Date').size().reset_index(name='Install Count')
            chart1 = alt.Chart(installs_over_time).mark_bar().encode(
                x='Date:T',
                y='Install Count:Q',
                tooltip=['Date', 'Install Count']
            ).properties(title='Installs Over Time')
            st.altair_chart(chart1, use_container_width=True)

            # 2. Installs per Technician
            installs_per_tech = filtered_df.groupby('Tech').size().reset_index(name='Install Count')
            chart2 = alt.Chart(installs_per_tech).mark_bar().encode(
                x='Tech:N',
                y='Install Count:Q',
                tooltip=['Tech', 'Install Count']
            ).properties(title='Installs per Technician')
            st.altair_chart(chart2, use_container_width=True)

            # 3. Drop Size Distribution
            drop_dist = filtered_df['Drop Size'].value_counts().reset_index()
            drop_dist.columns = ['Drop Size', 'Count']
            chart3 = alt.Chart(drop_dist).mark_bar().encode(
                x='Drop Size:N',
                y='Count:Q',
                tooltip=['Drop Size', 'Count']
            ).properties(title='Drop Size Distribution')
            st.altair_chart(chart3, use_container_width=True)

            # 4. Stacked bar: Drop Size per Tech
            stacked_df = filtered_df.groupby(['Tech', 'Drop Size']).size().reset_index(name='Count')
            chart4 = alt.Chart(stacked_df).mark_bar().encode(
                x='Tech:N',
                y='Count:Q',
                color='Drop Size:N',
                tooltip=['Tech', 'Drop Size', 'Count']
            ).properties(title='Drop Size by Technician')
            st.altair_chart(chart4, use_container_width=True)

            # 5. Line chart: Install trend over time
            line_chart = alt.Chart(installs_over_time).mark_line(point=True).encode(
                x='Date:T',
                y='Install Count:Q',
                tooltip=['Date', 'Install Count']
            ).properties(title='Install Trend Over Time')
            st.altair_chart(line_chart, use_container_width=True)

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Upload an Excel file to begin.")
