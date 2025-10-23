import streamlit as st
import pandas as pd
import tempfile
import os
from Script_Q2_v3_1 import read_excel_file, extract_unique_markets, create_market_sheets

# Set page configuration
st.set_page_config(page_title="RIA Template Builder - Unit Explanation Generator", layout="wide")

# Title and description
st.title("RIA Template Builder - Unit Explanation Generator")
st.markdown("""
This application is designed to generate explanations by unit, built by the **Template Builder RIA team in Madrid**.

### Instructions:
1. Upload an Excel file that follows the input format shown in the example file `Q3_input.xlsx`.
2. The app will process the file using the backend logic and generate a downloadable Excel file.
3. Ensure the sheet name is `DBQ2` and the column `Business Unit` is present.

**Note:** The input format is critical for successful processing. Please refer to the example file for guidance.
""")

# File upload
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

# Process file if uploaded
if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_file.read())
        tmp_input_path = tmp_input.name

    # Define parameters
    sheet_name = "DBQ2"
    market_column = "Business Unit"

    # Read and process the file
    df = read_excel_file(tmp_input_path, sheet_name)
    if df is not None and market_column in df.columns:
        unique_markets = extract_unique_markets(df, market_column)

        # Create output file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:
            output_path = tmp_output.name

        create_market_sheets(df, unique_markets, market_column, output_path)

        # Provide download link
        with open(output_path, "rb") as f:
            st.success("Processing complete! Download your file below:")
            st.download_button(label="Download Processed Excel File", data=f, file_name="Processed_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("The uploaded file does not match the expected format. Please ensure it follows the structure of Q3_input.xlsx.")