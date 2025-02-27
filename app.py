import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO
from docx import Document
import requests

def load_library():
    library_file = "Library_data.xlsx"
    if os.path.exists(library_file):
        df = pd.read_csv(library_file) if library_file.endswith('.csv') else pd.read_excel(library_file, engine='openpyxl')
        df.columns = [col.strip() for col in df.columns if col is not None]
        return df
    else:
        st.error("Library data file 'Library_data.xlsx' is missing. Please upload a valid library file.")
        return None

def load_master_data():
    master_file = "Muuto_Master_Data_CON_January_2025_EUR.xlsx"
    if os.path.exists(master_file):
        df = pd.read_excel(master_file, engine='openpyxl')
        df.columns = [col.strip() for col in df.columns if col is not None]
        return df
    else:
        st.error("Master data file is missing. Please upload a valid master file.")
        return None

# Indlæs Library-data og Master-data én gang
Library_data = load_library()
Master_data = load_master_data()

def load_excel(file):
    try:
        if file is None:
            raise ValueError("File not provided")
        excel_data = pd.ExcelFile(file, engine='openpyxl')
        return {sheet: pd.read_excel(excel_data, sheet_name=sheet) for sheet in excel_data.sheet_names}
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def clean_column_names(df):
    df.columns = df.iloc[1].astype(str).str.strip().str.lower().str.replace(' ', ' ')
    df = df[2:].reset_index(drop=True)
    
    # Ensure correct mapping of 'Article No.' and 'Quantity'
    column_mapping = {}
    for col in df.columns:
        if col.lower().strip() in ["article no.", "item variant number", "item no.", "article number"]:
            column_mapping[col] = "Article No."
        elif col.lower().strip() == "quantity":
            column_mapping[col] = "Quantity"
    
    df.rename(columns=column_mapping, inplace=True)
    
    st.write("Cleaned and mapped User Data Columns:", df.columns.tolist())
    return df
    df.columns = df.iloc[1].astype(str).str.strip().str.lower().str.replace(' ', ' ')
    df = df[2:].reset_index(drop=True)
    st.write("Cleaned User Data Columns:", df.columns.tolist())
    return df
    df.columns = df.iloc[1].astype(str).str.strip()
    return df[2:].reset_index(drop=True)

def match_columns(user_df):
    st.write("Debug - User Data Columns:", user_df.columns.tolist())  # Debugging log
    possible_columns = ["Article No.", "Item variant number", "Item no.", "Article Number"]
    match_column = next((col for col in user_df.columns if col.lower().strip() in [pc.lower().strip() for pc in possible_columns]), None)
    if match_column is None:
        st.error("The uploaded file must contain one of the expected item number columns: 'Item variant number', 'Item no.', 'Article No.', or 'Article Number'.")
        st.stop()
    return match_column
    possible_columns = ["Article No.", "Item variant number", "Item no."]
    match_column = next((col for col in user_df.columns if col.lower() in [pc.lower() for pc in possible_columns]), None)
    if match_column is None:
        st.error("The uploaded file must contain either 'Item variant number', 'Item no.', or 'Article No.'")
        st.stop()
    return match_column

def merge_library_data(user_df, library_df):
    match_column = match_columns(user_df)
    required_columns = ['EUR item no.', 'Product']
    for col in required_columns:
        if col not in library_df.columns:
            st.error(f"Column '{col}' not found in Library_data. Available columns: {library_df.columns}")
            st.stop()
    merged_df = user_df.merge(library_df[['EUR item no.', 'Product']], left_on=match_column, right_on='EUR item no.', how='left')
    merged_df['Output'] = merged_df['Quantity'].astype(str) + ' X ' + merged_df['Product'].fillna('Unknown')
    return merged_df[['Output']]

def generate_order_import_file(user_df):
    if 'Quantity' not in user_df.columns or 'Article No.' not in user_df.columns:
        st.error("Required columns not found in uploaded file. Check column names and header row.")
        st.stop()
    order_data = user_df[['Quantity', 'Article No.']].copy()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        order_data.to_excel(writer, index=False, header=False)
    buffer.seek(0)
    return buffer

def generate_sku_mapping(user_df, library_df, master_df):
    # Debugging: Print column names before merging
    st.write("User Data Columns:", user_df.columns.tolist())
    st.write("Library Data Columns:", library_df.columns.tolist())
    st.write("Master Data Columns:", master_df.columns.tolist())
    
    # Ensure columns are stripped of spaces
    user_df.columns = user_df.columns.str.strip().str.lower().str.replace(' ', ' ')
    library_df.columns = library_df.columns.str.strip().str.lower().str.replace(' ', ' ')
    master_df.columns = master_df.columns.str.strip()
    
    # Validate necessary columns exist
    if 'Article No.' not in user_df.columns:
        st.error("Column 'Article No.' not found in uploaded file.")
        st.stop()
    if 'EUR item no.' not in library_df.columns:
        st.error("Column 'EUR item no.' not found in Library Data.")
        st.stop()
    if 'ITEM NO.' not in master_df.columns:
        st.error("Column 'ITEM NO.' not found in Master Data.")
        st.stop()
    
    # Proceed with merging
    mapping = user_df.merge(library_df, left_on='Article No.', right_on='EUR item no.', how='left')
    mapping = mapping[['Quantity', 'Product', 'EUR item no.', 'GBP item no.', 'APMEA item no.', 'USD pattern no.']]
    
    master_data = user_df.merge(master_df, left_on='Article No.', right_on='ITEM NO.', how='left')
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        mapping.to_excel(writer, sheet_name='Item number mapping', index=False)
        master_data.to_excel(writer, sheet_name='Masterdata', index=False)
    buffer.seek(0)
    return buffer
    mapping = user_df.merge(library_df, left_on='Article No.', right_on='EUR item no.', how='left')
    mapping = mapping[['Quantity', 'Product', 'EUR item no.', 'GBP item no.', 'APMEA item no.', 'USD pattern no.']]
    
    master_data = user_df.merge(master_df, left_on='Article No.', right_on='ITEM NO.', how='left')
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        mapping.to_excel(writer, sheet_name='Item number mapping', index=False)
        master_data.to_excel(writer, sheet_name='Masterdata', index=False)
    buffer.seek(0)
    return buffer

st.title('Muuto Product List Generator')

st.write("""
This tool is designed to **help you structure, validate, and enrich pCon product data effortlessly**.
### **How it works:**  
1. **Upload your product list** – Export it from pCon as an Excel file.  
2. **Automated data matching** – The tool cross-references your data with Muuto’s official product library and master data.  
3. **Download structured files** – Choose from three ready-to-use formats:  
   - **Product list for presentations** – A clean list to support sales and visual presentations.  
   - **Order import file** – A structured file for seamless order uploads to the partner platform.  
   - **SKU mapping & master data** – A detailed overview linking item numbers to relevant product details.  

[Download an example file](https://raw.githubusercontent.com/TinaMuuto/Master-Converter/f280308cf9991b7eecb63e44ecac52dfb49482cf/pCon%20-%20exceleksport.xlsx)
""")

uploaded_file = st.file_uploader("Upload your product list (Excel or CSV)", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None and Library_data is not None and Master_data is not None:
    if uploaded_file.name.endswith(".csv"):
        user_df = pd.read_csv(uploaded_file)
    else:
        user_df = load_excel(uploaded_file)
        if 'Article List' in user_df:
            user_df = clean_column_names(user_df['Article List'])
        else:
            st.error("No 'Article List' sheet found in the uploaded file.")
            st.stop()
    
    if st.button("Download product list for presentations"):
        merged_df = merge_library_data(user_df, Library_data)
        buffer = BytesIO()
        doc = Document()
        doc.add_heading('Product List for Presentations', level=1)
        for row in merged_df['Output']:
            doc.add_paragraph(row)
        doc.save(buffer)
        buffer.seek(0)
        st.download_button("Download product list for presentations", buffer, file_name="product-list_presentation.docx")
    
    if st.button("Download product list for order import in partner platform"):
        buffer = generate_order_import_file(user_df)
        st.download_button("Download order import file", buffer, file_name="order-import.xlsx")
    
    if st.button("Download masterdata and SKU mapping"):
        buffer = generate_sku_mapping(user_df, Library_data, Master_data)
        st.download_button("Download SKU mapping", buffer, file_name="masterdata-SKUmapping.xlsx")
