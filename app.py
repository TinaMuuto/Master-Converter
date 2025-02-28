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
            # Attempt reading with semicolon separator first
            try:
                return pd.read_csv(uploaded_file, sep=';', engine='python')
            except pd.errors.ParserError:
                # If that fails, attempt comma separator
                return pd.read_csv(uploaded_file, sep=',', engine='python')
        else:
            return pd.ExcelFile(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None


def preprocess_user_data(df):
    # Skip the first row and reset index
    df = df.iloc[1:].reset_index(drop=True)

    # Extract relevant columns from the known structure
    df['Article No.'] = df.iloc[:, 17].astype(str).str.upper()  # Column R
    df['Quantity'] = df.iloc[:, 30]  # Column AE
    df['Variant'] = df.iloc[:, 4].fillna('').str.upper()  # Column E
    df['Short text'] = df.iloc[:, 2].fillna('').str.upper()  # Column C
    df['Base Article No.'] = df['Article No.'].str.split('-').str[0].str.upper()

    return df[['Article No.', 'Quantity', 'Variant', 'Short text', 'Base Article No.']]


def generate_word_file(merged_df):
    buffer = BytesIO()
    doc = Document()
    doc.add_heading('Product List for Presentations', level=1)

    # Indsæt hver linje fra kolonnen 'Word Output' i et nyt afsnit
    for row in merged_df['Word Output']:
        doc.add_paragraph(row)

    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generate_excel_file(merged_df, include_headers=True):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, header=include_headers)
    buffer.seek(0)
    return buffer


def match_article_numbers(user_df, master_df, library_df):
    # Sørg for, at de krævede kolonner findes i Master Data
    required_master_cols = ['ITEM NO.', 'PRODUCT']
    for col in required_master_cols:
        if col not in master_df.columns:
            st.error(f"Master Data file is missing required column: '{col}'")
            return pd.DataFrame()

    # Sørg for, at de krævede kolonner findes i Library Data
    required_library_cols = ['EUR ITEM NO.', 'PRODUCT']
    for col in required_library_cols:
        if col not in library_df.columns:
            st.error(f"Library Data file is missing required column: '{col}'")
            return pd.DataFrame()

    # Gør kolonnenavne ensartede for at sikre match
    master_df['ITEM NO.'] = master_df['ITEM NO.'].astype(str).str.upper()
    library_df['EUR ITEM NO.'] = library_df['EUR ITEM NO.'].astype(str).str.upper()

    # Forsøg først et direkte match i Master Data
    merged_df = user_df.merge(
        master_df[['ITEM NO.', 'PRODUCT']],
        left_on='Article No.',
        right_on='ITEM NO.',
        how='left'
    )

    # Hvis der ikke er noget match, forsøg at matche i Library Data
    unmatched = merged_df['PRODUCT'].isna()
    library_match = user_df.loc[unmatched].merge(
        library_df[['EUR ITEM NO.', 'PRODUCT']],
        left_on='Article No.',
        right_on='EUR ITEM NO.',
        how='left'
    )
    merged_df.loc[unmatched, 'PRODUCT'] = library_match['PRODUCT']

    # Stadig unmatched? Forsøg fallback på Base Article No. i Master Data
    unmatched = merged_df['PRODUCT'].isna()
    fallback_df = user_df.loc[unmatched].merge(
        master_df[['ITEM NO.', 'PRODUCT']],
        left_on='Base Article No.',
        right_on='ITEM NO.',
        how='left'
    )
    # Bevar original variant fra user_df
    merged_df.loc[unmatched, 'PRODUCT'] = fallback_df['PRODUCT']
    merged_df.loc[unmatched, 'Variant'] = user_df.loc[unmatched, 'Variant']

    # Forsøg samme fallback i Library Data
    library_fallback = user_df.loc[unmatched].merge(
        library_df[['EUR ITEM NO.', 'PRODUCT']],
        left_on='Base Article No.',
        right_on='EUR ITEM NO.',
        how='left'
    )
    merged_df.loc[unmatched, 'PRODUCT'] = library_fallback['PRODUCT']
    merged_df.loc[unmatched, 'Variant'] = user_df.loc[unmatched, 'Variant']

    # FINAL VARIANT
    merged_df['FINAL VARIANT'] = merged_df.apply(
        lambda row: row['Variant'] if row['Variant'] not in ['', 'LIGHT OPTION: OFF'] else row['Short text'],
        axis=1
    )

    # Opret Masterdata Output-kolonnen
    merged_df['Masterdata Output'] = (
        merged_df['Base Article No.'].fillna('') + " - " + merged_df['FINAL VARIANT'].fillna('')
    ).str.upper()

    # Opret Word Output-kolonnen
    merged_df['Word Output'] = merged_df.apply(
        lambda row: (
            f"{row['Quantity']} X {row['PRODUCT']} " +
            (f"- {row['FINAL VARIANT']}" if row['FINAL VARIANT'] not in ['', 'LIGHT OPTION: OFF'] else '')
        ) if pd.notna(row['PRODUCT']) else (
            f"{row['Quantity']} X {row['Short text']} " +
            (f"- {row['FINAL VARIANT']}" if row['FINAL VARIANT'] not in ['', 'LIGHT OPTION: OFF'] else '')
        ),
        axis=1
    ).str.upper()

    return merged_df[
        ['Quantity', 'Article No.', 'PRODUCT', 'Masterdata Output', 'Word Output']
    ]


# Streamlit-app titel og introduktion
st.title('Muuto Product List Generator')

st.write("""
This tool is designed to **help you structure, validate, and enrich pCon product data effortlessly**.

### **How it works:**  
1. **Export your product list from pCon** (formatted like the example file).  
2. **Upload your pCon file** to the app.  
3. **Click one of the three buttons** to generate the file you need.  
4. **Once generated, a new button will appear** for you to download the file.  
""")

# Fil-upload
uploaded_file = st.file_uploader("Upload your product list (Excel or CSV)", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    # Hent master- og biblioteksdata
    master_data = load_data("Muuto_Master_Data_CON_January_2025_EUR.xlsx")
    library_data = load_data("Library_data.xlsx")

    user_data = load_uploaded_file(uploaded_file)

    # Tjek om Excel-filen har et ark ved navn 'Article List'
    if isinstance(user_data, pd.ExcelFile) and 'Article List' in user_data.sheet_names:
        uploaded_df = pd.read_excel(user_data, sheet_name='Article List')
    else:
        uploaded_df = user_data

    # Sørg for, at vi har valid data
    if uploaded_df is not None:
        # Forbehandl brugerens data
        user_df = preprocess_user_data(uploaded_df)

        # Match artikelnumre
        matched_df = match_article_numbers(user_df, master_data, library_data)

        # Knap: Generer Word-fil til præsentationer
        if st.button("Generate product list for presentations"):
            buffer = generate_word_file(matched_df)
            st.download_button(
                "Download file", buffer, file_name="product-list_presentation.docx"
            )

        # Knap: Generer Excel-fil til ordreimport
        if st.button("Generate order import file"):
            buffer = generate_excel_file(
                matched_df[['Quantity', 'Article No.']], include_headers=False
            )
            st.download_button("Download file", buffer, file_name="order-import.xlsx")

        # Knap: Generer Masterdata + SKU Mapping
        if st.button("Generate masterdata and SKU mapping"):
            buffer = generate_excel_file(
                matched_df[['Quantity', 'Article No.', 'Masterdata Output']]
            )
            st.download_button("Download file", buffer, file_name="masterdata-SKUmapping.xlsx")
