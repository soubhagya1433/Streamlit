import streamlit as st
import pandas as pd
import xlwings as xw
import tempfile
import os

def open_excel_with_xlwings(file_path):
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open(file_path)
    return app, wb

st.title("ATA XlFile Roundtrip")
st.sidebar.image("C:/Anaconda/DE Project/ATA phase 1/pages/ATAlogo.png", use_column_width=True)
uploaded_file = st.file_uploader("Choose an Excel file to view and Edit in Excel UI", type="xlsx", key="file_uploader")

if uploaded_file is not None:
    # Create a persistent file path in session state
    if 'file_path' not in st.session_state or st.session_state.uploaded_file != uploaded_file:
        # Remove old file references
        if 'file_path' in st.session_state:
            os.unlink(st.session_state.file_path)  # Remove the old file

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getvalue())
            st.session_state.file_path = tmp.name
        
        # Update uploaded file in session state
        st.session_state.uploaded_file = uploaded_file

    if st.button("Open in Excel"):
        app, wb = open_excel_with_xlwings(st.session_state.file_path)
        st.session_state.excel_app = app
        st.session_state.excel_wb = wb
        st.session_state.excel_open = True
        st.success("Excel file is now open. You can make changes directly in Excel.")
        
    if os.path.exists(st.session_state.file_path):
        df = pd.read_excel(st.session_state.file_path)
        st.write("Displayed Data (may need refresh after changes):")
        st.dataframe(df)

    if st.button("Refresh Data"):
        if os.path.exists(st.session_state.file_path):
            df = pd.read_excel(st.session_state.file_path, engine='openpyxl')
            #st.write("Updated Data:")
            #st.dataframe(df)
        else:
            st.error("File does not exist.")

    file_name_input = st.text_input("Enter Excel filename for download (without extension):", "modified_excel_file")

    if st.button("Create File"):
        if file_name_input:
            # Ensure the filename ends with .xlsx
            file_name = f"{file_name_input}.xlsx"

            # Check if the file exists
            if os.path.exists(st.session_state.file_path):
                with open(st.session_state.file_path, "rb") as file:
                    file_data = file.read()
                # Create a download button that triggers the download immediately
                btn = st.download_button(
                    label="Download Excel File",
                    data=file_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                if btn:
                    st.success("File ready to download.")
            else:
                st.error("File does not exist.")
        else:
            st.error("Please provide a filename before downloading.")
