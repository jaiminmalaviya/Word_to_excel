import streamlit as st
import pandas as pd
from docx import Document
import os
import uuid
import time

def read_word_file(file_path):
    try:
        document = Document(file_path)
        data = []
        record = []
        for paragraph in document.paragraphs:
            line = paragraph.text.strip()
            if line:
                record.append(line)
                if len(record) == num_columns:
                    data.append(record)
                    record = []
        return data
    except Exception as e:
        st.error(f"Error reading the Word file: {e}")

def save_to_excel(data, output_file):
    try:
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(output_file, index=False)
        st.success(f"Excel file '{output_file}' created successfully.")
    except Exception as e:
        st.error(f"Error saving to Excel file: {e}")

def delete_file(file_path):
    try:
        os.remove(file_path)
    except Exception as e:
        st.error(f"Error deleting file '{file_path}': {e}")

if __name__ == "__main__":
    st.title("Word to Excel Converter")
    uploaded_file = st.file_uploader("Upload a Word file", type=["docx"])

    if uploaded_file is not None:
        file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": uploaded_file.size}
        num_columns = st.number_input("Enter the number of columns:", min_value=1, max_value=10, value=6)
        default_columns = ['Name', 'Position', 'Location', 'Industry', 'Phone', 'Email']
        column_names = st.text_input("Enter column names separated by commas (default: Name, Position, Location, Industry, Phone, Email):", value=", ".join(default_columns))
        columns = [col.strip() for col in column_names.split(",")]

        if len(columns) != num_columns:
            st.warning(f"Number of columns provided ({len(columns)}) does not match the specified number ({num_columns}). Using default column names.")
            columns = default_columns

        unique_id = uuid.uuid4().hex[:7]  # Generate a unique ID
        output_file = f"EMAIL LIST ({unique_id}).xlsx"  # Append the ID to the output file name
        data = read_word_file(uploaded_file)
        if data:
            save_to_excel(data, output_file)
            # time.sleep(3)  # Add a short delay
                        
            # Provide a download link for the file
            with open(output_file, "rb") as file:
                file_contents = file.read()
            st.download_button(label="Download Excel file", data=file_contents, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            delete_file(output_file)  # Delete the file after generating download link