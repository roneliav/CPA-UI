import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import base64
from pdf2image import convert_from_bytes
from openpyxl import load_workbook

# --- Page Configuration ---
st.set_page_config(
    page_title="CPA Office Utils",
    page_icon="�",
    layout="centered"
)

# Constants
API_ENDPOINT = "https://cpa-api.vercel.app/api/"
# API_ENDPOINT = "http://localhost:3000/api/"


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

def get_pdf_base64(pdf_path):
    with open(pdf_path, "rb") as f:
        pdf_data = f.read()
    return base64.b64encode(pdf_data).decode("utf-8")

def get_visible_sheets(file):
    try:
        # reset pointer and try to detect visible sheets via openpyxl (works for .xlsx)
        file.seek(0)
        wb = load_workbook(filename=file, read_only=True, data_only=True)
        return [ws.title for ws in wb.worksheets if ws.sheet_state == 'visible']
    except Exception:
        # If detection fails (e.g. .xls or missing library), fall back to including all sheets
        return None

def merge_files(files):
    # Create an empty list to store individual DataFrames
    df_list = []
    
    # Loop through each uploaded file
    for file in files:
        try:
            visible_sheets = get_visible_sheets(file)
            file.seek(0)
            excel_dfs = pd.read_excel(file, sheet_name=None, header=None)

            for sheet_name, df in excel_dfs.items():
                df = df.dropna(axis=0, how='all')
                # Only add the sheet if it's visible (or if we couldn't detect visibility)
                if visible_sheets is None or sheet_name in visible_sheets:
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


def get_file_info(file_content):
    file_content = file.read()
    file.seek(0)  # Reset file pointer for potential reuse
    base64.b64encode(file_content).decode('utf-8')


def get_base64_string(img):
    buffered = BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode('utf-8')

def get_file_info(file): 
    file_content = file.read() 
    file.seek(0) # Reset file pointer for potential reuse 
    b64_str = base64.b64encode(file_content).decode('utf-8') # If the file is a PDF and the base64 string is empty (or nearly empty), 
    # convert it to an image and return the base64 of the first image. 
    if not b64_str.strip(): 
        images = convert_from_bytes(file_content, dpi=200) 
        if images: 
            b64_str = [get_base64_string(img) for img in images]
            return b64_str, 'image'
    return b64_str, 'pdf'

# 2. Processing and Download Logic
if uploaded_files:
    st.info(f"✅ הועלו {len(uploaded_files)} קבצים בהצלחה")
    
    # Add send button to trigger processing
    if st.button("שלח"):
        try:
            if option == "איחוד קבצי אקסל":
                with st.spinner('מעבד קבצי אקסל...'):
                    merged_df = merge_files(uploaded_files)
            else:
                with st.spinner('מעבד קבצי PDF...'):
                    # Prepare files for API request
                    files = []
                    for file in uploaded_files:
                        file_base64, file_type = get_file_info(file)
                        files.append({
                            "file_name": file.name,
                            "file_base64": file_base64,
                            "file_type": file_type
                        })
                    
                    # Send files to external API
                    response = requests.post(
                        f"{API_ENDPOINT}", # extract_from_pdf
                        json={"files": files}
                    )

                    if response.status_code == 200:
                        print("response:", response)
                        # response_data = response.json()
                        # merged_df = pd.DataFrame(response_data['data'])
                    else:
                        error_detail = response.json().get('detail', response.text)
                        st.error(f"שגיאה בעיבוד הקבצים: {error_detail}")
                        st.stop()
            
            # Display results if processing was successful
            st.success("✅ העיבוד הושלם בהצלחה!")
            
            st.subheader("תצוגה מקדימה של הנתונים")
            st.dataframe(merged_df.head())

            # Convert the merged DataFrame to an Excel file in memory
            excel_data = to_excel(merged_df)

            st.subheader("הורדת הקובץ")
            st.download_button(
                label="📥 הורד קובץ אקסל",
                data=excel_data,
                file_name="extracted_data.xlsx" if option == "חילוץ טבלאות מ-PDF" else "merged_files.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"שגיאה בעיבוד הקבצים: {str(e)}")

else:
    file_type = "אקסל" if option == "איחוד קבצי אקסל" else "PDF"
    st.warning(f"☝️ אנא העלה לפחות קובץ {file_type} אחד כדי להתחיל")