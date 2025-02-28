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
    Forventer:
        Kolonne A: Product
        Kolonne B: EUR item no. (unikt)
        Kolonne C: GBP item no.
        Kolonne D: APMEA item no.
        Kolonne E: USD pattern no.
        Kolonne F: Match Status
    """
    if not os.path.exists(library_path):
        st.error(f"Filen {library_path} mangler i mappen. Upload eller placér filen korrekt.")
        return None
    
    try:
        df = pd.read_excel(library_path, engine="openpyxl")
        # Rens kolonner for mellemrum og store bogstaver:
        df.columns = df.columns.str.strip().str.upper()
        # Sørg også for at selve kolonneindhold i EUR ITEM NO. er i uppercase (til match):
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
        # Rens kolonner for mellemrum og sæt dem i uppercase:
        df.columns = df.columns.str.strip().str.upper()
        # Gør ITEM NO. uppercase
        if "ITEM NO." in df.columns:
            df["ITEM NO."] = df["ITEM NO."].astype(str).str.strip().str.upper()
        return df
    except Exception as e:
        st.error(f"Fejl ved indlæsning af {master_path}: {e}")
        return None


def load_user_file(uploaded_file):
    """
    Læser brugerens uploadede fil.
    - Hvis Excel: led efter en fane "Article List"
    - Hvis CSV: forsøg først at læse med semikolon, hvis fejl, brug komma
    Returnerer en DataFrame eller None.
    """
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith(".csv"):
            # 1) Første forsøg: sep=';'
            try:
                df = pd.read_csv(uploaded_file, sep=';', engine="python", header=0)
            except:
                # 2) Hvis fejl, forsøg med sep=','
                uploaded_file.seek(0)  # nulstil filpegeren
                df = pd.read_csv(uploaded_file, sep=',', engine="python", header=0)
            return df
        
        else:
            # Excel
            excel_obj = pd.ExcelFile(uploaded_file, engine="openpyxl")
            if "Article List" in excel_obj.sheet_names:
                df = pd.read_excel(excel_obj, sheet_name="Article List", header=0)
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
    Henter kolonnerne:
      - index 17 (Article No.)
      - index 30 (Quantity)
      - index 2  (Short text)
      - index 4  (Variant text)
    Data starter fra række 2 => df.iloc[1:] 
    """
    # Fjern første række, hvis overskrifter ligger på row 2 i arket
    df = df.iloc[1:].reset_index(drop=True)

    # Tjek at df har nok kolonner
    num_cols = df.shape[1]
    if num_cols < 31:
        st.error("Den uploadede fil indeholder ikke nok kolonner (mindst 31 kræves). Tjek format.")
        return None

    # Uddrag de nødvendige kolonner:
    article_no = df.iloc[:, 17].astype(str).str.strip().str.upper()
    quantity = df.iloc[:, 30]
    short_text = df.iloc[:, 2].astype(str).str.strip().str.upper()
    variant = df.iloc[:, 4].astype(str).str.strip().str.upper()

    # Læg dem i et nyt DataFrame
    out_df = pd.DataFrame({
        "ARTICLE_NO": article_no,
        "QUANTITY": quantity,
        "SHORT_TEXT": short_text,
        "VARIANT": variant
    })

    return out_df


#####################
# 3. PRÆSENTATIONSLISTE (WORD)
#####################

def generate_presentation_word(df_user, df_library):
    """
    1) For hver række i brugerdata:
       - Forsøg at matche 'ARTICLE_NO' i df_library['EUR ITEM NO.']
       - Hvis match: "QUANTITY X PRODUCT"
       - Hvis intet match: "QUANTITY X SHORT_TEXT - VARIANT" (men udelad '- VARIANT' hvis variant er tom eller 'LIGHT OPTION: OFF')
    2) Generér Word-fil i hukommelsen og returnér en BytesIO
    """

    # Tjek at library har kolonner: PRODUCT, EUR ITEM NO.
    required_cols = ["PRODUCT", "EUR ITEM NO."]
    for col in required_cols:
        if col not in df_library.columns:
            st.error(f"Library_data mangler kolonne: {col}")
            return None

    # Lav opslagsdict
    lookup_library = df_library.set_index("EUR ITEM NO.")["PRODUCT"].to_dict()

    # Opret liste til Word-tekster
    word_lines = []

    for _, row in df_user.iterrows():
        article_no = row["ARTICLE_NO"]
        quantity = row["QUANTITY"]
        short_text = row["SHORT_TEXT"]
        variant = row["VARIANT"]

        product_match = lookup_library.get(article_no)
        if not product_match:
            short_article = article_no.split('-')[0].strip().upper()
            product_match = lookup_library.get(short_article)

        if product_match:
            text_line = f"{quantity} X {product_match}"
        else:
            # fallback
            if variant and variant != "LIGHT OPTION: OFF":
                text_line = f"{quantity} X {short_text} - {variant}"
            else:
                text_line = f"{quantity} X {short_text}"

        # Alt i uppercase
        text_line = text_line.upper()
        word_lines.append(text_line)

    # Opret Word-dokument i hukommelsen
    buffer = BytesIO()
    doc = Document()
    doc.add_heading('Product List for Presentations', level=1)
    for line in word_lines:
        doc.add_paragraph(line)
    doc.save(buffer)
    buffer.seek(0)
    return buffer


#####################
# 4. ORDER IMPORT FIL (2 kolonner, ingen header)
#####################

def generate_order_import_excel(df_user):
    """
    Returnerer en BytesIO med 2 kolonner, ingen headers:
      - col A: QUANTITY
      - col B: ARTICLE_NO
    Filnavn: order-import.xlsx
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
    1) Opret sheet "Item number mapping":
       - For hver ARTICLE_NO i df_user => forsøg match i df_library['EUR ITEM NO.']
         - direct match
         - fallback: splitted match
       - Returner:
         A: Quantity in setting
         B: Product in setting
         C: EUR item no.
         D: GBP item no.
         E: APMEA item no.
         F: USD pattern no.
         G: Match status
       
    2) Opret sheet "Master data export":
       - For hver ARTICLE_NO i df_user => forsøg match i df_master['ITEM NO.']
         - direct match
         - fallback: splitted match
       - Returner samtlige kolonner fra df_master for de matchede rækker
    """

    # Tjek at library har de nødvendige kolonner
    required_lib_cols = ["PRODUCT", "EUR ITEM NO.", "GBP ITEM NO.", "APMEA ITEM NO.", "USD PATTERN NO.", "MATCH STATUS"]
    for col in required_lib_cols:
        if col not in df_library.columns:
            st.error(f"Library_data mangler kolonnen {col} til SKU mapping.")
            return None

    # Omdøb kolonner i library til et konsistent sæt
    df_library_renamed = df_library.rename(columns={
        "PRODUCT": "LIB_PRODUCT",
        "EUR ITEM NO.": "LIB_EUR_ITEM_NO",
        "GBP ITEM NO.": "LIB_GBP_ITEM_NO",
        "APMEA ITEM NO.": "LIB_APMEA_ITEM_NO",
        "USD PATTERN NO.": "LIB_USD_PATTERN_NO",
        "MATCH STATUS": "LIB_MATCH_STATUS"
    })

    # Step 1: direct match
    merged_direct = pd.merge(
        df_user,
        df_library_renamed,
        how="left",
        left_on="ARTICLE_NO",
        right_on="LIB_EUR_ITEM_NO"
    )

    # Identificer hvem der ikke fik match
    mask_no_direct = merged_direct["LIB_PRODUCT"].isna()

    # Fallback
    fallback_rows = merged_direct.loc[mask_no_direct].copy()
    fallback_rows["SHORT_ARTICLE"] = fallback_rows["ARTICLE_NO"].str.split('-').str[0].str.strip().str.upper()

    fallback_merged = pd.merge(
        fallback_rows.drop(columns=["LIB_PRODUCT", "LIB_EUR_ITEM_NO", "LIB_GBP_ITEM_NO", "LIB_APMEA_ITEM_NO", "LIB_USD_PATTERN_NO", "LIB_MATCH_STATUS"]),
        df_library_renamed,
        how="left",
        left_on="SHORT_ARTICLE",
        right_on="LIB_EUR_ITEM_NO"
    )

    for col in ["LIB_PRODUCT", "LIB_EUR_ITEM_NO", "LIB_GBP_ITEM_NO", "LIB_APMEA_ITEM_NO", "LIB_USD_PATTERN_NO", "LIB_MATCH_STATUS"]:
        merged_direct.loc[mask_no_direct, col] = fallback_merged[col].values

    # Byg "Item number mapping"-DataFrame
    item_number_mapping_df = pd.DataFrame({
        "Quantity in setting": merged_direct["QUANTITY"],
        "Product in setting": merged_direct["LIB_PRODUCT"],
        "EUR item no.": merged_direct["LIB_EUR_ITEM_NO"],
        "GBP item no.": merged_direct["LIB_GBP_ITEM_NO"],
        "APMEA item no.": merged_direct["LIB_APMEA_ITEM_NO"],
        "USD pattern no.": merged_direct["LIB_USD_PATTERN_NO"],
        "Match status": merged_direct["LIB_MATCH_STATUS"]
    })

    # --- Forbered "Master data export" ---
    if "ITEM NO." not in df_master.columns:
        st.error("Masterdata-filen mangler kolonnen 'ITEM NO.'")
        return None

    # Saml unikke article_no fra df_user
    user_article_list = df_user["ARTICLE_NO"].unique()
    all_master_matches = pd.DataFrame()

    for article in user_article_list:
        # direct match
        direct_matches = df_master[df_master["ITEM NO."] == article]
        if len(direct_matches) == 0:
            short_article = article.split("-")[0].strip().upper()
            direct_matches = df_master[df_master["ITEM NO."] == short_article]

        if len(direct_matches) > 0:
            all_master_matches = pd.concat([all_master_matches, direct_matches], ignore_index=True)

    # Fjern evt. duplicates
    all_master_matches.drop_duplicates(inplace=True)

    # Skriv 2 sheets til én fil
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        item_number_mapping_df.to_excel(writer, sheet_name="Item number mapping", index=False)
        all_master_matches.to_excel(writer, sheet_name="Master data export", index=False)
    buffer.seek(0)
    return buffer


#####################
# 6. Streamlit-app
#####################

def main():
    st.set_page_config(page_title="Muuto Product List Generator", layout="centered")

    # Vise instruktioner (som i briefet)
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

    # Indlæs referencefiler (Library_data og Master_data)
    df_library = load_library_data()
    df_master = load_master_data()
    if (df_library is None) or (df_master is None):
        return  # Stop, hvis de ikke kunne indlæses

    # Filupload fra bruger
    uploaded_file = st.file_uploader("Upload your product list (Excel or CSV)", type=['xlsx', 'xls', 'csv'])

    if uploaded_file:
        df_user_raw = load_user_file(uploaded_file)
        if df_user_raw is not None:
            # Preprocess
            df_user = preprocess_user_data(df_user_raw)
            if df_user is None:
                return  # Kun hvis preprocess fejlede
            
            # 1) Word til præsentationer
            if st.button("Generate List for presentations"):
                word_buffer = generate_presentation_word(df_user, df_library)
                if word_buffer:
                    st.download_button("Download Word file", data=word_buffer, file_name="product-list.docx")

            # 2) Ordreimport
            if st.button("Generate product list for order import in partner platform"):
                order_buffer = generate_order_import_excel(df_user)
                st.download_button("Download Excel file", data=order_buffer, file_name="order-import.xlsx")

            # 3) SKU mapping & masterdata
            if st.button("Generate SKU mapping & masterdata"):
                sku_buffer = generate_sku_masterdata_excel(df_user, df_library, df_master)
                if sku_buffer:
                    st.download_button("Download Excel file", data=sku_buffer, file_name="SKUmapping-masterdata.xlsx")


if __name__ == "__main__":
    main()
