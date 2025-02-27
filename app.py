import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from docx import Document

def load_excel(file):
    try:
        return pd.read_excel(file, sheet_name=None, engine='openpyxl')
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def clean_column_names(df):
    df.columns = df.iloc[1].astype(str).str.strip()
    return df[2:].reset_index(drop=True)

def merge_library_data(user_df, library_df):
    merged_df = user_df.merge(library_df[['EUR item no.', 'Product']], left_on='Article No.', right_on='EUR item no.', how='left')
    merged_df['Output'] = merged_df['Quantity'].astype(str) + ' X ' + merged_df['Product'].fillna('Unknown')
    return merged_df[['Output']]

def generate_presentation_doc(merged_df):
    doc = Document()
    doc.add_heading('Product List for Presentations', level=1)
    for row in merged_df['Output']:
        doc.add_paragraph(row)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_order_import_file(user_df):
    order_data = user_df[['Quantity', 'Article No.']].copy()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        order_data.to_excel(writer, index=False, header=False)
    buffer.seek(0)
    return buffer

def generate_sku_mapping(user_df, library_df, master_df):
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

[Download an example file](https://raw.githubusercontent.com/TinaMuuto/Master-Converter/main/pCon%20-%20exceleksport.xlsx)
""")

uploaded_file = st.file_uploader("Upload your product list (Excel)", type=['xlsx', 'xls'])

if uploaded_file:
    user_data = load_excel(uploaded_file)
    library_data = load_excel("/mnt/data/Library_data.xlsx")["Sheet1"]
    master_data = load_excel("/mnt/data/Muuto_Master_Data_CON_January_2025_EUR.xlsx")["Sheet1"]
    
    if 'Article List' in user_data:
        user_df = clean_column_names(user_data['Article List'])
    else:
        st.error("No 'Article List' sheet found in the uploaded file.")
        st.stop()
    
    if st.button("Download product list for presentations"):
        merged_df = merge_library_data(user_df, library_data)
        buffer = generate_presentation_doc(merged_df)
        st.download_button("Download product list for presentations", buffer, file_name="product-list_presentation.docx")
    
    if st.button("Download product list for order import in partner platform"):
        buffer = generate_order_import_file(user_df)
        st.download_button("Download order import file", buffer, file_name="order-import.xlsx")
    
    if st.button("Download masterdata and SKU mapping"):
        buffer = generate_sku_mapping(user_df, library_data, master_data)
        st.download_button("Download SKU mapping", buffer, file_name="masterdata-SKUmapping.xlsx")
