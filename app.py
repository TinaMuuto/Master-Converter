import streamlit as st
import pandas as pd
import os
from io import BytesIO
from utils import load_library_data, match_item_numbers, generate_product_list_presentation, generate_order_import, generate_detailed_product_list, generate_masterdata

# Filnavne for opslag
LIBRARY_FILE = "Library_data.xlsx"
MASTER_FILE = "Muuto_Master_Data_CON_January_2025_EUR.xlsx"

# Load library data
library_df = load_library_data(LIBRARY_FILE)
master_df = load_library_data(MASTER_FILE)

# Streamlit UI
st.title("Muuto Product List Generator")
st.write("This app allows you to upload an Excel or CSV file and generate various product lists by matching with Muuto's reference data.")

st.subheader("How it works")
st.markdown("""
1. Upload your file (Excel or CSV).
2. The app automatically identifies the relevant column for lookup.
3. It matches your data against Muuto’s library.
4. You can download four different files:
   - Product list for presentations (Word)
   - Order import file (Excel)
   - Detailed product list with item number mapping (Excel)
   - Master data for products in setting (Excel)
""")

uploaded_file = st.file_uploader("Upload your file", type=["csv", "xlsx", "xlsm", "xls"])

if uploaded_file:
    user_df = pd.read_excel(uploaded_file, None) if uploaded_file.name.endswith(("xlsx", "xlsm", "xls")) else {"Sheet1": pd.read_csv(uploaded_file, delimiter=";", encoding="utf-8")}
    
    # Identificer lookup-kolonne
    lookup_column = match_item_numbers(user_df)

    if lookup_column:
        st.success(f"Identified lookup column: {lookup_column}")

        # Generer outputfiler
        product_list_docx = generate_product_list_presentation(user_df, lookup_column, library_df)
        order_import_excel = generate_order_import(user_df, lookup_column)
        detailed_product_excel = generate_detailed_product_list(user_df, lookup_column, library_df)
        masterdata_excel = generate_masterdata(user_df, lookup_column, library_df, master_df)

        # Download knapper
        st.download_button("Download product list for presentations", data=product_list_docx, file_name="product-list_presentation.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.download_button("Download product list for order import in partner platform", data=order_import_excel, file_name="order-import.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Download detailed product list with item number mapping", data=detailed_product_excel, file_name="detailed-product-list.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Download masterdata for products in setting", data=masterdata_excel, file_name="masterdata.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("Could not find a matching column in the uploaded file.")

