import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO
from docx import Document

def load_data(file_path):
    if os.path.exists(file_path):
        return pd.read_excel(file_path, engine='openpyxl')
    else:
        st.error(f"File {file_path} is missing. Please upload a valid file.")
        return None

def preprocess_user_data(df):
    df = df.iloc[1:].reset_index(drop=True)  # Skip header row
    df['Article No.'] = df.iloc[:, 17].astype(str)  # Column R
    df['Quantity'] = df.iloc[:, 30]  # Column AE
    df['Description'] = df.iloc[:, 4]  # Column E
    df['Base Article No.'] = df['Article No.'].str.split('-').str[0]
    return df[['Article No.', 'Quantity', 'Description', 'Base Article No.']]

def match_article_numbers(user_df, master_df):
    master_df['ITEM NO.'] = master_df['ITEM NO.'].astype(str)
    
    # Exact match first
    merged_df = user_df.merge(
        master_df[['ITEM NO.', 'PRODUCT DESCRIPTION']], 
        left_on='Article No.', 
        right_on='ITEM NO.', 
        how='left'
    )
    
    # If no exact match, try matching on base article number
    unmatched = merged_df['PRODUCT DESCRIPTION'].isna()
    fallback_df = user_df[unmatched].merge(
        master_df[['ITEM NO.', 'PRODUCT DESCRIPTION']], 
        left_on='Base Article No.', 
        right_on='ITEM NO.', 
        how='left'
    )
    
    # Combine exact matches and fallback matches
    merged_df.loc[unmatched, 'PRODUCT DESCRIPTION'] = fallback_df['PRODUCT DESCRIPTION']
    
    # Create final description
    merged_df['Final Description'] = merged_df.apply(
        lambda row: f"{row['PRODUCT DESCRIPTION']} - {row['Description']}" if pd.notna(row['PRODUCT DESCRIPTION']) else "Unknown",
        axis=1
    )
    
    return merged_df[['Quantity', 'Article No.', 'Final Description']]

def generate_word_file(merged_df):
    buffer = BytesIO()
    doc = Document()
    doc.add_heading('Product List for Presentations', level=1)
    for row in merged_df['Final Description']:
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

uploaded_file = st.file_uploader("Upload your product list (Excel)", type=['xlsx'])
if uploaded_file and master_data is not None:
    user_data = pd.ExcelFile(uploaded_file, engine='openpyxl')
    if 'Article List' in user_data.sheet_names:
        user_df = preprocess_user_data(pd.read_excel(user_data, sheet_name='Article List'))
        matched_df = match_article_numbers(user_df, master_data)
        
        if st.button("Generate product list for presentations"):
            buffer = generate_word_file(matched_df)
            st.download_button("Download file", buffer, file_name="product-list_presentation.docx")
        
        if st.button("Generate order import file"):
            buffer = generate_excel_file(matched_df[['Quantity', 'Article No.']])
            st.download_button("Download file", buffer, file_name="order-import.xlsx")
        
        if st.button("Generate masterdata and SKU mapping"):
            buffer = generate_excel_file(matched_df)
            st.download_button("Download file", buffer, file_name="masterdata-SKUmapping.xlsx")
