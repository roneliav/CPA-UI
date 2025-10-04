import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="CPA Office Utils",
    page_icon="�",
    layout="centered"
)

# Constants
API_ENDPOINT = "https://cpa-api.vercel.app/api/"

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
        
        # Set the worksheet to right-to-left
        worksheet = writer.sheets['MergedData']
        worksheet.right_to_left()

    processed_data = output.getvalue()
    return processed_data


def merge_files(files):
    # Create an empty list to store individual DataFrames
    df_list = []
    
    # Loop through each uploaded file
    for file in files:
        try:
            # Read the Excel file into a DataFrame
            excel_dfs = pd.read_excel(file, sheet_name=None, header=None)
            for sheet_name, df in excel_dfs.items():
                df = df.dropna(axis=0, how='all')
                # Add a source column to track which file the data came from
                # df['מקור'] = file.name
                df_list.append(df)
                
        except Exception as e:
            st.error(f"שגיאה בקריאת הקובץ {file.name}: {e}")

    # Ensure we have something to merge
    if df_list:
        # Merge all DataFrames in the list into a single DataFrame
        return pd.concat(df_list, ignore_index=True)
    st.error(f"שגיאה: לא נמצאו קבצים לאחד.")



# --- App UI ---
st.title("� CPA Office Utils")

# Add selection for different functionalities
option = st.selectbox(
    "בחר פעולה",
    ["איחוד קבצי אקסל", "חילוץ טבלאות מ-PDF"],
    format_func=lambda x: x
)

if option == "איחוד קבצי אקסל":
    st.subheader("📄 איחוד קבצי אקסל")
    st.write(
        "העלה קובץ אקסל אחד או יותר עם מבנה עמודות זהה. "
        "המערכת תאחד אותם לקובץ אחד להורדה."
    )
else:
    st.subheader("📑 חילוץ טבלאות מ-PDF")
    st.write(
        "העלה קובץ PDF אחד או יותר. "
        "המערכת תחלץ את הטבלאות ותייצא אותן לקובץ אקסל."
    )

# Enable a button to remove all uploaded files
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 1

# 1. File Uploader
if option == "איחוד קבצי אקסל":
    uploaded_files = st.file_uploader(
        "בחר קבצי אקסל (.xlsx, .xls)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key=st.session_state['uploader_key']
    )
else:
    uploaded_files = st.file_uploader(
        "בחר קבצי PDF",
        type=['pdf'],
        accept_multiple_files=True,
        key=st.session_state['uploader_key']
    )


if st.button("Clear all files"):
    # st.session_state.uploaded_files = []
    st.session_state["uploader_key"] += 1
    st.rerun()   # Refresh the app so uploader resets



# 2. Processing and Download Logic
if uploaded_files:
    st.info(f"✅ הועלו {len(uploaded_files)} קבצים בהצלחה")
    
    try:
        if option == "איחוד קבצי אקסל":
            merged_df = merge_files(uploaded_files)

        else:
            with st.spinner('מעבד קבצי PDF...'):
                # Prepare files for API request
                files = [
                    ('files', (file.name, file.getvalue(), 'application/pdf'))
                    for file in uploaded_files
                ]
                
                # Send files to external API
                response = requests.post(f"{API_ENDPOINT}/extract-tables", files=files)
                df = response['df']
                
                if response.status_code == 200:
                    # Get the Excel file from the response
                    response_data = response.json()
                    excel_bytes = bytes(response_data['data'])
                    
                    # Convert bytes to DataFrame for preview
                    excel_buffer = BytesIO(excel_bytes)
                    merged_df = pd.read_excel(excel_buffer)
                else:
                    error_detail = response.json().get('detail', response.text)
                    st.error(f"שגיאה בעיבוד הקבצים: {error_detail}")
                    st.stop()



        st.subheader("תצוגה מקדימה של הנתונים המאוחדים")
        st.dataframe(merged_df.head())

        # Convert the merged DataFrame to an Excel file in memory
        excel_data = to_excel(merged_df)

        st.subheader("הורדת הקובץ המאוחד")
        st.download_button(
            label="📥 הורד קובץ אקסל מאוחד",
            data=excel_data,
            file_name="merged_files.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"שגיאה בעיבוד הקבצים: {str(e)}")

else:
    file_type = "אקסל" if option == "איחוד קבצי אקסל" else "PDF"
    st.warning(f"☝️ אנא העלה לפחות קובץ {file_type} אחד כדי להתחיל")