import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO
from docx import Document

def load_data(file_path):
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, engine='openpyxl', index_col=None)
        df.columns = [col.strip().upper() for col in df.columns]  # Normalize column names to uppercase
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
            return pd.ExcelFile(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def preprocess_user_data(df):
    df = df.iloc[1:].reset_index(drop=True)  # Start from row 2 (zero-indexed)
    df['Article No.'] = df.iloc[:, 17].astype(str).str.upper()  # Column R
    df['Quantity'] = df.iloc[:, 30]  # Column AE
    df['Variant'] = df.iloc[:, 4].fillna('').str.upper()  # Column E
    df['Short text'] = df.iloc[:, 2].fillna('').str.upper()  # Column C
    df['Base Article No.'] = df['Article No.'].str.split('-').str[0].str.upper()  # Base article number for fallback
    return df[['Article No.', 'Quantity', 'Variant', 'Short text', 'Base Article No.']]

def match_article_numbers(user_df, master_df, library_df):
    required_master_cols = ['ITEM NO.', 'PRODUCT']
    required_library_cols = ['EUR ITEM NO.', 'PRODUCT']
    
    for col in required_master_cols:
        if col not in master_df.columns:
            st.error(f"Master Data file is missing required column: '{col}'")
            return pd.DataFrame()
    
    for col in required_library_cols:
        if col not in library_df.columns:
            st.error(f"Library Data file is missing required column: '{col}'")
            return pd.DataFrame()
    
    master_df['ITEM NO.'] = master_df['ITEM NO.'].astype(str).str.upper()
    library_df['EUR ITEM NO.'] = library_df['EUR ITEM NO.'].astype(str).str.upper()
    
    # Exact match in Master Data
    merged_df = user_df.merge(
        master_df[['ITEM NO.', 'PRODUCT']], 
        left_on='Article No.', 
        right_on='ITEM NO.', 
        how='left'
    )
    
    # Exact match in Library Data if no match in Master Data
    unmatched = merged_df['PRODUCT'].isna()
    library_match = user_df[unmatched].merge(
        library_df[['EUR ITEM NO.', 'PRODUCT']],
        left_on='Article No.', 
        right_on='EUR ITEM NO.', 
        how='left'
    )
    merged_df.loc[unmatched, 'PRODUCT'] = library_match['PRODUCT']
    
    # If no match, find the closest match using Base Article No.
    unmatched = merged_df['PRODUCT'].isna()
    fallback_df = user_df[unmatched].merge(
        master_df[['ITEM NO.', 'PRODUCT']], 
        left_on='Base Article No.', 
        right_on='ITEM NO.', 
        how='left'
    )
    merged_df.loc[unmatched, 'PRODUCT'] = fallback_df['PRODUCT']
    
    library_fallback = user_df[unmatched].merge(
        library_df[['EUR ITEM NO.', 'PRODUCT']],
        left_on='Base Article No.', 
        right_on='EUR ITEM NO.', 
        how='left'
    )
    merged_df.loc[unmatched, 'PRODUCT'] = library_fallback['PRODUCT']
    
    # Ensure correct variant handling when no exact match is found
    merged_df['FINAL VARIANT'] = merged_df.apply(
        lambda row: row['Variant'] if row['Variant'] not in ['', 'LIGHT OPTION: OFF'] else row['Short text'], axis=1
    )
    
    # If still no match, adjust based on output type
    merged_df['Masterdata Output'] = (merged_df['Base Article No.'].fillna('') + " - " + merged_df['FINAL VARIANT'].fillna('')).str.upper()
    merged_df['Word Output'] = merged_df.apply(
        lambda row: f"{row['Quantity']} X {row['PRODUCT']} {' - ' + row['FINAL VARIANT'] if row['FINAL VARIANT'] not in ['', 'LIGHT OPTION: OFF'] else ''}"
        if pd.notna(row['PRODUCT']) else
        f"{row['Quantity']} X {row['Short text']} {' - ' + row['FINAL VARIANT'] if row['FINAL VARIANT'] not in ['', 'LIGHT OPTION: OFF'] else ''}", axis=1
    ).str.upper()
    
    return merged_df[['Quantity', 'Article No.', 'PRODUCT', 'Masterdata Output', 'Word Output']]

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
            buffer = generate_excel_file(matched_df[['Quantity', 'Article No.']], include_headers=False)
            st.download_button("Download file", buffer, file_name="order-import.xlsx")
        
        if st.button("Generate masterdata and SKU mapping"):
            buffer = generate_excel_file(matched_df[['Quantity', 'Article No.', 'Masterdata Output']])
            st.download_button("Download file", buffer, file_name="masterdata-SKUmapping.xlsx")
