
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Dashboard", layout="wide")
st.title("📊 Excel File Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # Load all sheets
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        st.sidebar.header("Available Sheets")
        selected_sheet = st.sidebar.selectbox("Choose a sheet", sheet_names)

        df = pd.read_excel(xls, sheet_name=selected_sheet)
        st.subheader(f"Data Preview: {selected_sheet}")
        st.dataframe(df, use_container_width=True)

        # Show basic stats if numeric data present
        numeric_cols = df.select_dtypes(include="number").columns
        if not numeric_cols.empty:
            st.subheader("Summary Statistics")
            st.write(df[numeric_cols].describe())

            chart_col = st.selectbox("Choose a column for charting", numeric_cols)
            chart_type = st.selectbox("Choose chart type", ["Bar", "Line", "Area"])

            if chart_type == "Bar":
                st.bar_chart(df[chart_col])
            elif chart_type == "Line":
                st.line_chart(df[chart_col])
            elif chart_type == "Area":
                st.area_chart(df[chart_col])

    except Exception as e:
        st.error(f"Error reading file: {e}")
else:
    st.info("Upload an Excel file to begin.")
