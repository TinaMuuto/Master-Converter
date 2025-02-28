import streamlit as st
import pandas as pd
import openpyxl
import os
from io import BytesIO
from docx import Document

#####################
# 1. Data-load-funktioner
#####################

def load_library_data(library_path="Library_data.xlsx"):
    """
    Indlæser Library_data.xlsx, renser kolonnenavne og returnerer en DataFrame.
    Forventer (i række 1, evt. i blandet case):
        - Product
        - EUR item no.
        - GBP item no.
        - APMEA item no.
        - USD pattern no.
        - Match Status
    """
    if not os.path.exists(library_path):
        st.error(f"Filen {library_path} mangler i mappen. Upload eller placér filen korrekt.")
        return None
    
    try:
        df = pd.read_excel(library_path, engine="openpyxl")
        # Konverter kolonnenavne til uppercase + strip
        df.columns = df.columns.str.strip().str.upper()
        # Konverter 'EUR ITEM NO.' kolonner til uppercase indhold
        if "EUR ITEM NO." in df.columns:
            df["EUR ITEM NO."] = df["EUR ITEM NO."].astype(str).str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Fejl ved indlæsning af {library_path}: {e}")
        return None


def load_master_data(master_path="Muuto_Master_Data_CON_January_2025_EUR.xlsx"):
    """
    Indlæser hele masterdata-filen.
    Forventer kolonne B: ITEM NO. (unikt).
    Returnerer en DataFrame med alle kolonner.
    """
    if not os.path.exists(master_path):
        st.error(f"Filen {master_path} mangler i mappen. Upload eller placér filen korrekt.")
        return None
    
    try:
        df = pd.read_excel(master_path, engine="openpyxl")
        df.columns = df.columns.str.strip().str.upper()
        if "ITEM NO." in df.columns:
            df["ITEM NO."] = df["ITEM NO."].astype(str).str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Fejl ved indlæsning af {master_path}: {e}")
        return None


def load_user_file(uploaded_file):
    """
    Læser brugerens uploadede fil (pCon-export).
    - Hvis Excel: led efter en fane "Article List", skip de to første rækker, ingen header
    - Hvis CSV: forsøg først med sep=';', hvis fejl, forsøg med sep=','
      i begge tilfælde skiprows=2, header=None
    Returnerer en DataFrame (uden navngivne kolonner) eller None.
    """
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith(".csv"):
            # 1) Prøv først sep=';'
            try:
                df = pd.read_csv(uploaded_file, sep=';', engine="python", header=None, skiprows=2)
            except:
                # 2) Fallback: sep=','
                uploaded_file.seek(0)  # nulstil filpegeren
                df = pd.read_csv(uploaded_file, sep=',', engine="python", header=None, skiprows=2)
            return df
        
        else:
            # Excel
            excel_obj = pd.ExcelFile(uploaded_file, engine="openpyxl")
            if "Article List" in excel_obj.sheet_names:
                df = pd.read_excel(excel_obj, sheet_name="Article List", skiprows=2, header=None)
                return df
            else:
                st.error("Filen indeholder ikke en fane ved navn 'Article List'.")
                return None
    
    except Exception as e:
        st.error(f"Fejl ved læsning af fil: {e}")
        return None

#####################
# 2. Preprocessing af brugerens data
#####################

def preprocess_user_data(df):
    """
    Henter kolonnerne (index-baseret):
      - 17 -> Article No.
      - 30 -> Quantity
      - 2  -> Short Text
      - 4  -> Variant Text
    Antager, at skiprows=2 allerede har fjernet de to første rækker.
    """
    if df.shape[1] < 31:
        st.error("Den uploadede fil indeholder ikke nok kolonner (mindst 31 kræves). Tjek format.")
        return None

    article_no = df.iloc[:, 17].astype(str).str.strip().str.upper()
    quantity = df.iloc[:, 30]
    short_text = df.iloc[:, 2].astype(str).str.strip().str.upper()
    variant_text = df.iloc[:, 4].astype(str).str.strip().str.upper()

    out_df = pd.DataFrame({
        "ARTICLE_NO": article_no,
        "QUANTITY": quantity,
        "SHORT_TEXT": short_text,
        "VARIANT_TEXT": variant_text
    })
    return out_df


#####################
# 3. PRÆSENTATIONSLISTE (WORD)
#####################

def generate_presentation_word(df_user, df_library):
    """
    1) For hver række i df_user:
       - Forsøg direct match: 'ARTICLE_NO' i df_library['EUR ITEM NO.']
       - Hvis match: "QUANTITY X PRODUCT" (fuldt match - ingen partial)
       - Hvis intet match: "QUANTITY X SHORT_TEXT - VARIANT_TEXT" 
         (udelad '- VARIANT_TEXT' hvis variant_text er tom eller "LIGHT OPTION: OFF")
    2) Sortér efter (PRODUCT el. SHORT_TEXT) i stigende alfabetisk rækkefølge
    3) Returnér Word-fil som BytesIO
    """
    required_cols = ["PRODUCT", "EUR ITEM NO."]
    for col in required_cols:
        if col not in df_library.columns:
            st.error(f"Library_data mangler kolonnen: '{col}' for at generere præsentationslisten.")
            return None

    # Opslagsdict: {EUR ITEM NO.: PRODUCT}
    lookup_library = df_library.set_index("EUR ITEM NO.")["PRODUCT"].to_dict()

    lines_info = []
    for _, row in df_user.iterrows():
        article_no = row["ARTICLE_NO"]
        quantity = row["QUANTITY"]
        short_text = row["SHORT_TEXT"]
        variant_text = row["VARIANT_TEXT"]

        # Kun fuldt match
        product_match = lookup_library.get(article_no)

        if product_match:
            sort_key = product_match  # Sortér efter product
            final_line = f"{quantity} X {product_match}"
        else:
            sort_key = short_text     # Sortér efter short_text, da der ikke er match
            if variant_text and variant_text != "LIGHT OPTION: OFF":
                final_line = f"{quantity} X {short_text} - {variant_text}"
            else:
                final_line = f"{quantity} X {short_text}"

        # Vi gemmer (sort_key.upper(), final_line.upper()) så sorteres case-insensitivt
        lines_info.append((sort_key.upper(), final_line.upper()))

    # Sorter linjerne alfabetisk på sort_key
    lines_info.sort(key=lambda x: x[0])

    # Opret Word-dokument i hukommelsen
    buffer = BytesIO()
    doc = Document()
    doc.add_heading('Product List for Presentations', level=1)
    for _, line_text in lines_info:
        doc.add_paragraph(line_text)
    doc.save(buffer)
    buffer.seek(0)
    return buffer


#####################
# 4. ORDER IMPORT FIL (2 kolonner, ingen header)
#####################

def generate_order_import_excel(df_user):
    """
    Returnerer en BytesIO med 2 kolonner, uden headers:
      - A: QUANTITY
      - B: ARTICLE_NO
    """
    buffer = BytesIO()
    temp_df = df_user[["QUANTITY", "ARTICLE_NO"]].copy()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        temp_df.to_excel(writer, index=False, header=False)
    buffer.seek(0)
    return buffer


#####################
# 5. SKU MAPPING & MASTERDATA
#####################

def generate_sku_masterdata_excel(df_user, df_library, df_master):
    """
    Genererer 2 ark:
    1) "Item number mapping":
       - Forsøger direct + fallback match (split på '-')
       - Kolonner:
         Quantity in setting
         Article No.
         Short Text
         Variant text
         Product in setting
         EUR item no.
         GBP item no.
         APMEA item no.
         USD pattern no.
         Match status

    2) "Master data export":
       - Direct + fallback match
       - Returnerer alle master-kolonner + 'Article No.', 'Short Text', 'Variant text'
    """

    # =============== ITEM NUMBER MAPPING ===============
    # Tjek nødvendige kolonner
    lib_required = ["PRODUCT", "EUR ITEM NO.", "GBP ITEM NO.", 
                    "APMEA ITEM NO.", "USD PATTERN NO.", "MATCH STATUS"]
    # Lav en liste over evt. manglende kolonner:
    missing_cols = [c for c in lib_required if c not in df_library.columns]
    if missing_cols:
        st.warning(f"Følgende kolonner mangler i Library_data (ingen KeyError, men de vil ikke blive opdateret): {missing_cols}")

    # Omdøb kun dem, der reelt findes
    rename_map = {
        "PRODUCT": "LIB_PRODUCT",
        "EUR ITEM NO.": "LIB_EUR_ITEM_NO",
        "GBP ITEM NO.": "LIB_GBP_ITEM_NO",
        "APMEA ITEM NO.": "LIB_APMEA_ITEM_NO",
        "USD PATTERN NO.": "LIB_USD_PATTERN_NO",
        "MATCH STATUS": "LIB_MATCH_STATUS"
    }
    valid_renames = {}
    for old_col, new_col in rename_map.items():
        if old_col in df_library.columns:
            valid_renames[old_col] = new_col

    df_library_renamed = df_library.rename(columns=valid_renames)

    merged_direct = pd.merge(
        df_user,
        df_library_renamed,
        how="left",
        left_on="ARTICLE_NO",
        right_on="LIB_EUR_ITEM_NO" if "LIB_EUR_ITEM_NO" in df_library_renamed.columns else None
    )

    # fallback for item number mapping
    if "LIB_PRODUCT" in merged_direct.columns:
        mask_no_direct = merged_direct["LIB_PRODUCT"].isna()
    else:
        # Hvis "LIB_PRODUCT" ikke findes, er alt unmatched
        mask_no_direct = pd.Series([True]*len(merged_direct), index=merged_direct.index)

    fallback_rows = merged_direct.loc[mask_no_direct].copy()
    fallback_rows["SHORT_ARTICLE"] = fallback_rows["ARTICLE_NO"].str.split('-').str[0].str.strip().str.upper()

    fallback_merged = pd.merge(
        fallback_rows,
        df_library_renamed,
        how="left",
        left_on="SHORT_ARTICLE",
        right_on="LIB_EUR_ITEM_NO" if "LIB_EUR_ITEM_NO" in df_library_renamed.columns else None
    )

    # Her sætter vi kun kolonner, der reelt findes i fallback_merged
    columns_to_update = ["LIB_PRODUCT", "LIB_EUR_ITEM_NO", "LIB_GBP_ITEM_NO",
                         "LIB_APMEA_ITEM_NO", "LIB_USD_PATTERN_NO", "LIB_MATCH_STATUS"]
    for col in columns_to_update:
        if col in fallback_merged.columns:
            merged_direct.loc[mask_no_direct, col] = fallback_merged[col].values
        else:
            st.warning(f"Kolonnen '{col}' findes ikke i fallback_merged. Springer over for item number mapping.")

    # Byg item_number_mapping_df
    item_number_mapping_df = pd.DataFrame()
    item_number_mapping_df["Quantity in setting"] = merged_direct["QUANTITY"]
    item_number_mapping_df["Article No."] = merged_direct["ARTICLE_NO"]
    item_number_mapping_df["Short Text"] = merged_direct["SHORT_TEXT"]
    item_number_mapping_df["Variant text"] = merged_direct["VARIANT_TEXT"]
    item_number_mapping_df["Product in setting"] = merged_direct.get("LIB_PRODUCT", None)
    item_number_mapping_df["EUR item no."] = merged_direct.get("LIB_EUR_ITEM_NO", None)
    item_number_mapping_df["GBP item no."] = merged_direct.get("LIB_GBP_ITEM_NO", None)
    item_number_mapping_df["APMEA item no."] = merged_direct.get("LIB_APMEA_ITEM_NO", None)
    item_number_mapping_df["USD pattern no."] = merged_direct.get("LIB_USD_PATTERN_NO", None)
    item_number_mapping_df["Match status"] = merged_direct.get("LIB_MATCH_STATUS", None)


    # =============== MASTER DATA EXPORT ===============
    if "ITEM NO." not in df_master.columns:
        st.warning("Masterdata-filen mangler kolonnen 'ITEM NO.'. Ingen KeyError, men matches bliver tomme.")
        # Lav et tomt master_data_export_df
        master_data_export_df = pd.DataFrame(columns=["Article No.", "Short Text", "Variant text"] + df_master.columns.tolist())
    else:
        # Direct + fallback merging
        master_direct = pd.merge(
            df_user,
            df_master,
            how="left",
            left_on="ARTICLE_NO",
            right_on="ITEM NO."
        )

        mask_no_direct_master = master_direct["ITEM NO."].isna()
        fallback_master_rows = master_direct.loc[mask_no_direct_master].copy()
        fallback_master_rows["SHORT_ARTICLE"] = fallback_master_rows["ARTICLE_NO"].str.split('-').str[0].str.strip().str.upper()

        fallback_master_merged = pd.merge(
            fallback_master_rows,
            df_master,
            how="left",
            left_on="SHORT_ARTICLE",
            right_on="ITEM NO."
        )

        # Kopiér kun kolonner, der findes i fallback_master_merged
        for col in df_master.columns:
            if col in fallback_master_merged.columns:
                master_direct.loc[mask_no_direct_master, col] = fallback_master_merged[col].values
            else:
                st.warning(f"Kolonnen '{col}' findes ikke i fallback_master_merged. Springes over for master fallback.")

        master_direct.drop_duplicates(inplace=True)
        # Omdøb bruger-kolonner
        master_direct.rename(columns={
            "ARTICLE_NO": "Article No.",
            "SHORT_TEXT": "Short Text",
            "VARIANT_TEXT": "Variant text"
        }, inplace=True)

        # Flyt "Article No.", "Short Text", "Variant text" forrest
        front_cols = ["Article No.", "Short Text", "Variant text"]
        other_cols = [c for c in master_direct.columns if c not in front_cols]
        master_data_export_df = master_direct[front_cols + other_cols]

    # Skriv 2 ark i én Excel
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        item_number_mapping_df.to_excel(writer, sheet_name="Item number mapping", index=False)
        master_data_export_df.to_excel(writer, sheet_name="Master data export", index=False)

    buffer.seek(0)
    return buffer

#####################
# 6. Streamlit-app
#####################

def main():
    st.set_page_config(page_title="Muuto Product List Generator", layout="centered")

    # Instruktioner
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

    # Indlæs referencefiler
    df_library = load_library_data()
    df_master = load_master_data()
    if (df_library is None) or (df_master is None):
        return  # Stop, hvis de ikke kunne indlæses

    # Filupload
    uploaded_file = st.file_uploader("Upload your product list (Excel or CSV)", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        df_user_raw = load_user_file(uploaded_file)
        if df_user_raw is not None:
            # Preprocess
            df_user = preprocess_user_data(df_user_raw)
            if df_user is None:
                return  # Stop, hvis preprocess fejlede

            # Knap 1: Word-fil (ingen partial match)
            if st.button("Generate List for presentations"):
                word_buffer = generate_presentation_word(df_user, df_library)
                if word_buffer:
                    st.download_button(
                        "Download Word file", 
                        data=word_buffer, 
                        file_name="product-list.docx"
                    )

            # Knap 2: Order import (2 kolonner)
            if st.button("Generate product list for order import in partner platform"):
                order_buffer = generate_order_import_excel(df_user)
                st.download_button(
                    "Download Excel file", 
                    data=order_buffer, 
                    file_name="order-import.xlsx"
                )

            # Knap 3: SKU mapping & masterdata (med partial fallback)
            if st.button("Generate SKU mapping & masterdata"):
                sku_buffer = generate_sku_masterdata_excel(df_user, df_library, df_master)
                if sku_buffer:
                    st.download_button(
                        "Download Excel file", 
                        data=sku_buffer, 
                        file_name="SKUmapping-masterdata.xlsx"
                    )


if __name__ == "__main__":
    main()
