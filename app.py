import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO
from docx import Document

def load_data(file_path):
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = [col.strip() for col in df.columns]  # Strip whitespace from column names
        return df
    else:
        st.error(f"File {file_path} is missing. Please upload a valid file.")
        return None

def load_uploaded_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                return pd.read_csv(uploaded_file, sep=';', engine='python')
            except pd.errors.ParserError:
                return pd.read_csv(uploaded_file, sep=',', engine='python')
        else:
            df = pd.ExcelFile(uploaded_file, engine='openpyxl')
            df.sheet_names = [sheet.strip() for sheet in df.sheet_names]  # Strip sheet names
            return df
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def preprocess_user_data(df):
    df = df.iloc[1:].reset_index(drop=True)  # Start from row 2 (zero-indexed)
    df['Article No.'] = df.iloc[:, 17].astype(str)  # Column R
    df['Quantity'] = df.iloc[:, 30]  # Column AE
    df['Variant'] = df.iloc[:, 4]  # Column E
    df['Short text'] = df.iloc[:, 2]  # Column C
    df['Base Article No.'] = df['Article No.'].str.split('-').str[0]  # Base article number for fallback
    return df[['Article No.', 'Quantity', 'Variant', 'Short text', 'Base Article No.']]

def match_article_numbers(user_df, master_df, library_df):
    if not all(col in master_df.columns for col in ['ITEM NO.', 'Product']):
        st.error("Master Data file is missing required columns: 'ITEM NO.', 'Product'")
        return pd.DataFrame()
    
    if not all(col in library_df.columns for col in ['EUR item no.', 'Product']):
        st.error("Library Data file is missing required columns: 'EUR item no.', 'Product'")
        return pd.DataFrame()
    
    master_df['ITEM NO.'] = master_df['ITEM NO.'].astype(str)
    library_df['EUR item no.'] = library_df['EUR item no.'].astype(str)
    
    # Exact match in Master Data
    merged_df = user_df.merge(
        master_df[['ITEM NO.', 'Product']], 
        left_on='Article No.', 
        right_on='ITEM NO.', 
        how='left'
    )
    
    # Exact match in Library Data if no match in Master Data
    unmatched = merged_df['Product'].isna()
    library_match = user_df[unmatched].merge(
        library_df[['EUR item no.', 'Product']],
        left_on='Article No.', 
        right_on='EUR item no.', 
        how='left'
    )
    merged_df.loc[unmatched, 'Product'] = library_match['Product']
    
    # If no match, find the closest match using Base Article No.
    unmatched = merged_df['Product'].isna()
    fallback_df = user_df[unmatched].merge(
        master_df[['ITEM NO.', 'Product']], 
        left_on='Base Article No.', 
        right_on='ITEM NO.', 
        how='left'
    )
    merged_df.loc[unmatched, 'Product'] = fallback_df['Product']
    
    library_fallback = user_df[unmatched].merge(
        library_df[['EUR item no.', 'Product']],
        left_on='Base Article No.', 
        right_on='EUR item no.', 
        how='left'
    )
    merged_df.loc[unmatched, 'Product'] = library_fallback['Product']
    
    # If still no match, append Short text from Column C in the uploaded file
    merged_df.loc[merged_df['Product'].isna(), 'Product'] = merged_df['Base Article No.'] + " - " + merged_df['Short text']
    
    return merged_df[['Quantity', 'Article No.', 'Product']]

def generate_word_file(merged_df):
    buffer = BytesIO()
    doc = Document()
    doc.add_heading('Product List for Presentations', level=1)
    for row in merged_df['Product']:
        doc.add_paragraph(row)
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_excel_file(merged_df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, header=True)
    buffer.seek(0)
    return buffer

# Load master and library data
master_data = load_data("Muuto_Master_Data_CON_January_2025_EUR.xlsx")
library_data = load_data("Library_data.xlsx")

st.title('Muuto Product List Generator')

st.write("""
This tool is designed to **help you structure, validate, and enrich pCon product data effortlessly**.

### **How it works:**  
1. **Export your product list from pCon** (formatted like the example file).  
2. **Upload your pCon file** to the app.  
3. **Click one of the three buttons** to generate the file you need.  
4. **Once generated, a new button will appear** for you to download the file.  
""")

uploaded_file = st.file_uploader("Upload your product list (Excel or CSV)", type=['xlsx', 'csv'])
if uploaded_file and master_data is not None:
    user_data = load_uploaded_file(uploaded_file)
    if isinstance(user_data, pd.ExcelFile) and 'Article List' in user_data.sheet_names:
        uploaded_df = pd.read_excel(user_data, sheet_name='Article List')
    else:
        uploaded_df = user_data
    
    if uploaded_df is not None:
        user_df = preprocess_user_data(uploaded_df)
        matched_df = match_article_numbers(user_df, master_data, library_data)
        
        if st.button("Generate product list for presentations"):
            buffer = generate_word_file(matched_df)
            st.download_button("Download file", buffer, file_name="product-list_presentation.docx")
        
        if st.button("Generate order import file"):
            buffer = generate_excel_file(matched_df[['Quantity', 'Article No.']])
            st.download_button("Download file", buffer, file_name="order-import.xlsx")
        
        if st.button("Generate masterdata and SKU mapping"):
            buffer = generate_excel_file(matched_df)
            st.download_button("Download file", buffer, file_name="masterdata-SKUmapping.xlsx")
