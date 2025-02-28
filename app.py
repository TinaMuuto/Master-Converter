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
        df = pd.read_excel(library_file, engine='openpyxl')
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

def extract_fixed_columns(df):
   df = df.iloc[1:].reset_index(drop=True)  # Starter fra række 2
    article_no_col = df.iloc[:, 17]  # Kolonne R
    quantity_col = df.iloc[:, 30]  # Kolonne AE
    return pd.DataFrame({'Article No.': article_no_col, 'Quantity': quantity_col})

def merge_library_data(user_df, library_df):
    merged_df = user_df.merge(library_df[['EUR item no.', 'Product']], left_on='Article No.', right_on='EUR item no.', how='left')
    merged_df['Output'] = merged_df['Quantity'].astype(str) + ' X ' + merged_df['Product'].fillna('Unknown')
    return merged_df[['Output']]

def generate_order_import_file(user_df):
    order_data = user_df[['Quantity', 'Article No.']].copy()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        order_data.to_excel(writer, index=False, header=False)
    buffer.seek(0)
    return buffer

def generate_sku_mapping(user_df, library_df, master_df):
    mapping = user_df.merge(library_df[['EUR item no.', 'Product', 'GBP item no.', 'APMEA item no.', 'USD pattern no.']], left_on='Article No.', right_on='EUR item no.', how='left')
    master_data = user_df.merge(master_df, left_on='Article No.', right_on='ITEM NO.', how='left')
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        mapping.to_excel(writer, sheet_name='Item number mapping', index=False)
        master_data.to_excel(writer, sheet_name='Masterdata', index=False)
    buffer.seek(0)
    return buffer

import streamlit as st

st.title('Muuto Product List Generator')

st.write("""
This tool is designed to **help you structure, validate, and enrich pCon product data effortlessly**.

### **How it works:**  
1. **Export your product list from pCon** (formatted like the example file).  
2. **Upload your pCon file** to the app.  
3. **Click one of the three buttons** to generate the file you need.  
4. **Once generated, a new button will appear** for you to download the file.  

### **What can the app generate?**
#### 1. Product list for presentations
A Word file with product quantities and descriptions for easy copy-pasting into PowerPoint.

**Example output:**
- 1 X 70/70 Table / 170 X 85 CM / 67 X 33.5" - Solid Oak/Anthracite Black  
- 1 X Fiber Armchair / Swivel Base - Refine Leather Cognac/Anthracite Black  

#### 2. Product list for order import
A file formatted for direct import into the partner platform. This allows you to:
- Visualize the products  
- Place a quote/order  
- Pass the list to Customer Care to avoid manual entry  

#### 3. Product SKU mapping  
An Excel file with two sheets:
- **Product SKU mapping** – A list of products in the uploaded pCon setting with corresponding item numbers for EUR, UK, APMEA, and pattern numbers for the US.  
- **Master data export** – A full data export of the uploaded products for project documentation.  

[Download an example file](https://raw.githubusercontent.com/TinaMuuto/Master-Converter/f280308cf9991b7eecb63e44ecac52dfb49482cf/pCon%20-%20exceleksport.xlsx)
""")

uploaded_file = st.file_uploader("Upload your product list (Excel or CSV)", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None and Library_data is not None and Master_data is not None:
    user_df = load_excel(uploaded_file)
    if 'Article List' in user_df:
        user_df = extract_fixed_columns(user_df['Article List'])
    else:
        st.error("No 'Article List' sheet found in the uploaded file.")
        st.stop()
    
    if st.button("Generate product list for presentations"):
        merged_df = merge_library_data(user_df, Library_data)
        buffer = BytesIO()
        doc = Document()
        doc.add_heading('Product List for Presentations', level=1)
        for row in merged_df['Output']:
            doc.add_paragraph(row)
        doc.save(buffer)
        buffer.seek(0)
        st.download_button("Download file", buffer, file_name="product-list_presentation.docx")
    
    if st.button("Generate product list for order import in partner platform"):
        buffer = generate_order_import_file(user_df)
        st.download_button("Download file", buffer, file_name="order-import.xlsx")
    
    if st.button("Generate masterdata and SKU mapping"):
        buffer = generate_sku_mapping(user_df, Library_data, Master_data)
        st.download_button("Download file", buffer, file_name="masterdata-SKUmapping.xlsx")
