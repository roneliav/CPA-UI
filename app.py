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

def merge_files(files, mode='visible', selected_sheets=None):
    """
    Merge Excel files based on the specified mode.

    Args:
        files: List of uploaded files
        mode: 'all', 'visible', or 'selected'
        selected_sheets: Dict mapping file names to list of selected sheet names (for mode='selected')
    """
    # Create an empty list to store individual DataFrames
    df_list = []

    # Loop through each uploaded file
    for file in files:
        try:
            file.seek(0)
            excel_dfs = pd.read_excel(file, sheet_name=None, header=None)

            if mode == 'all':
                # Merge all sheets regardless of visibility
                for sheet_name, df in excel_dfs.items():
                    df = df.dropna(axis=0, how='all')
                    df_list.append(df)

            elif mode == 'visible':
                # Merge only visible sheets
                visible_sheets = get_visible_sheets(file)
                file.seek(0)
                for sheet_name, df in excel_dfs.items():
                    df = df.dropna(axis=0, how='all')
                    # Only add the sheet if it's visible (or if we couldn't detect visibility)
                    if visible_sheets is None or sheet_name in visible_sheets:
                        df_list.append(df)

            elif mode == 'selected':
                # Merge only selected sheets
                if selected_sheets and file.name in selected_sheets:
                    for sheet_name, df in excel_dfs.items():
                        if sheet_name in selected_sheets[file.name]:
                            df = df.dropna(axis=0, how='all')
                            df_list.append(df)

        except Exception as e:
            st.error(f"שגיאה בקריאת הקובץ {file.name}: {e}")

    # Ensure we have something to merge
    if df_list:
        # Merge all DataFrames in the list into a single DataFrame
        return pd.concat(df_list, ignore_index=True)
    st.error(f"שגיאה: לא נמצאו קבצים לאחד.")

def get_all_sheets_from_files(files):
    """
    Get all sheet names from uploaded files, grouped by file name.

    Returns:
        Dict mapping file names to list of sheet names
    """
    file_sheets = {}
    for file in files:
        try:
            file.seek(0)
            excel_file = pd.ExcelFile(file)
            file_sheets[file.name] = excel_file.sheet_names
        except Exception as e:
            st.error(f"שגיאה בקריאת הקובץ {file.name}: {e}")
    return file_sheets



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
        "בחר קבצי אקסל (.xlsx, .xls, .xlsm)",
        type=['xlsx', 'xls', 'xlsm'],
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

# Initialize session state for sheet selection
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = {}
if 'show_sheet_selection' not in st.session_state:
    st.session_state.show_sheet_selection = False

# 2. Processing and Download Logic
if uploaded_files:
    st.info(f"✅ הועלו {len(uploaded_files)} קבצים בהצלחה")

    # Add three buttons for Excel merge options
    if option == "איחוד קבצי אקסל":
        st.write("בחר אופן איחוד:")
        col1, col2, col3 = st.columns(3)

        with col1:
            merge_all_clicked = st.button("איחוד כל הגיליונות", use_container_width=True)
        with col2:
            merge_visible_clicked = st.button("איחוד גיליונות נראים", use_container_width=True)
        with col3:
            choose_sheets_clicked = st.button("בחירת גיליונות", use_container_width=True)

        # Handle "Choose sheets" button
        if choose_sheets_clicked:
            st.session_state.show_sheet_selection = True
            st.session_state.selected_sheets = {}

        # Display sheet selection UI if "Choose sheets" was clicked
        if st.session_state.show_sheet_selection:
            st.subheader("בחר גיליונות למיזוג")

            # Add custom CSS for green selected buttons
            st.markdown("""
                <style>
                .stButton > button[kind="primary"] {
                    background-color: #28a745 !important;
                    border-color: #28a745 !important;
                    color: white !important;
                }
                .stButton > button[kind="primary"]:hover {
                    background-color: #218838 !important;
                    border-color: #1e7e34 !important;
                }
                </style>
            """, unsafe_allow_html=True)

            file_sheets = get_all_sheets_from_files(uploaded_files)

            # Display sheets grouped by file
            for file_name, sheet_names in file_sheets.items():
                st.write(f"**{file_name}**")

                # Initialize selected sheets for this file if not exists
                if file_name not in st.session_state.selected_sheets:
                    st.session_state.selected_sheets[file_name] = []

                # Create buttons for each sheet
                cols = st.columns(min(4, len(sheet_names)))
                for idx, sheet_name in enumerate(sheet_names):
                    col_idx = idx % 4
                    with cols[col_idx]:
                        is_selected = sheet_name in st.session_state.selected_sheets[file_name]
                        button_label = f"✓ {sheet_name}" if is_selected else sheet_name
                        button_type = "primary" if is_selected else "secondary"

                        if st.button(button_label, key=f"{file_name}_{sheet_name}", type=button_type, use_container_width=True):
                            # Toggle sheet selection
                            if is_selected:
                                st.session_state.selected_sheets[file_name].remove(sheet_name)
                            else:
                                st.session_state.selected_sheets[file_name].append(sheet_name)
                            st.rerun()
                st.write("---")

            # Add merge button for selected sheets
            if any(sheets for sheets in st.session_state.selected_sheets.values()):
                if st.button("איחוד גיליונות נבחרים", type="primary"):
                    merge_selected_clicked = True
                else:
                    merge_selected_clicked = False
            else:
                st.warning("אנא בחר לפחות גיליון אחד")
                merge_selected_clicked = False
        else:
            merge_selected_clicked = False

        # Process based on which button was clicked
        if merge_all_clicked or merge_visible_clicked or merge_selected_clicked:
            try:
                with st.spinner('מעבד קבצי אקסל...'):
                    if merge_all_clicked:
                        merged_df = merge_files(uploaded_files, mode='all')
                    elif merge_visible_clicked:
                        merged_df = merge_files(uploaded_files, mode='visible')
                    elif merge_selected_clicked:
                        merged_df = merge_files(uploaded_files, mode='selected', selected_sheets=st.session_state.selected_sheets)
                        st.session_state.show_sheet_selection = False  # Hide selection UI after merge

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
                    file_name="merged_files.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"שגיאה בעיבוד הקבצים: {str(e)}")

    # Add send button for PDF extraction
    else:
        if st.button("שלח"):
            try:
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