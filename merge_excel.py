import streamlit as st
import pandas as pd
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="Excel File Merger",
    page_icon="📄",
    layout="centered"
)

# --- Helper Function for Excel Conversion ---
# This function converts a DataFrame to an in-memory Excel file.
# Using a function helps keep the main code clean and is good practice.
@st.cache_data
def to_excel(df):
    """Converts a DataFrame to an Excel file in memory."""
    output = BytesIO()
    # Use the 'xlsxwriter' engine for better compatibility.
    # index=False prevents pandas from writing row indices to the file.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='MergedData')
    processed_data = output.getvalue()
    return processed_data

# --- App UI ---
st.title("📄 Excel File Merger")
st.write(
    "Upload one or more Excel files with the same column structure. "
    "The app will merge them into a single downloadable file."
)

# 1. File Uploader
uploaded_files = st.file_uploader(
    "Choose Excel files (.xlsx)",
    type=['xlsx'],
    accept_multiple_files=True
)




# 2. Processing and Download Logic
# This block runs only when the user has uploaded at least one file.
if uploaded_files:
    st.info(f"✅ {len(uploaded_files)} file(s) uploaded successfully.")
    
    # Create an empty list to store individual DataFrames
    df_list = []
    
    # Loop through each uploaded file
    for file in uploaded_files:
        try:
            # Read the Excel file into a DataFrame
            excel_dfs = pd.read_excel(file, sheet_name=None)
            for sheet_name, df in excel_dfs.items():
                df = df.dropna(axis=0, how='all')

                # Add a source column to track which file the data came from
                # df['Source_File'] = filename        
                df_list.append(df)
                
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

    # Ensure we have something to merge
    if df_list:
        # Merge all DataFrames in the list into a single DataFrame
        # ignore_index=True resets the index of the merged DataFrame
        merged_df = pd.concat(df_list, ignore_index=True)
        
        st.subheader("Preview of Merged Data")
        st.dataframe(merged_df.head()) # Show the first 5 rows of the merged data

        # Convert the merged DataFrame to an Excel file in memory
        excel_data = to_excel(merged_df)

        st.subheader("Download Your Merged File")
        # 3. Download Button
        st.download_button(
            label="📥 Download Merged Excel File",
            data=excel_data,
            file_name="merged_files.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("☝️ Please upload at least one Excel file to get started.")