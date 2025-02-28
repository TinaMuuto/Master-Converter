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
        df.columns = df.columns.str.strip().str.upper()
        # Sørg for at kolonner bruges i uppercase
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
    Læser brugerens uploadede fil.
    - Hvis Excel: led efter en fane "Article List", skip de to første rækker, ingen header
    - Hvis CSV: forsøg først med sep=';', hvis fejl, forsøg med sep=',', i begge tilfælde skiprows=2, header=None
    Returnerer en DataFrame (uden navngivne kolonner) eller None.
    """
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith(".csv"):
            # Prøv først sep=';'
            try:
                df = pd.read_csv(uploaded_file, sep=';', engine="python", header=None, skiprows=2)
            except:
                # Fallback sep=','
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
    # Tjek at df har nok kolonner
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
       - Hvis match: "QUANTITY X PRODUCT"
       - Hvis intet match: "QUANTITY X SHORT_TEXT - VARIANT_TEXT"
         (udelad '- VARIANT_TEXT' hvis variant_text er tom eller "LIGHT OPTION: OFF")
    2) Returnér Word-fil som BytesIO
    """
    required_cols = ["PRODUCT", "EUR ITEM NO."]
    for col in required_cols:
        if col not in df_library.columns:
            st.error(f"Library_data mangler kolonne: {col}")
            return None

    # Opslagsdict
    lookup_library = df_library.set_index("EUR ITEM NO.")["PRODUCT"].to_dict()

    word_lines = []
    for _, row in df_user.iterrows():
        article_no = row["ARTICLE_NO"]
        quantity = row["QUANTITY"]
        short_text = row["SHORT_TEXT"]
        variant_text = row["VARIANT_TEXT"]

        product_match = lookup_library.get(article_no)
        if not product_match:
            short_article = article_no.split('-')[0].strip().upper()
            product_match = lookup_library.get(short_article)

        if product_match:
            text_line = f"{quantity} X {product_match}"
        else:
            # fallback
            if variant_text and variant_text != "LIGHT OPTION: OFF":
                text_line = f"{quantity} X {short_text} - {variant_text}"
            else:
                text_line = f"{quantity} X {short_text}"

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
       Kolonner:
         - Quantity in setting
         - Article No.
         - Short Text
         - Variant text
         - Product in setting
         - EUR item no.
         - GBP item no.
         - APMEA item no.
         - USD pattern no.
         - Match status

    2) "Master data export":
       - Indeholder ALLE kolonner fra df_master for de matchede rækker
         + 'Article No.', 'Short Text', 'Variant text' fra brugerfilen 
         (så man kan se, hvad der blev matchet).
    """

    # =============== 1) ITEM NUMBER MAPPING ===============
    required_lib_cols = ["PRODUCT", "EUR ITEM NO.", "GBP ITEM NO.", 
                         "APMEA ITEM NO.", "USD PATTERN NO.", "MATCH STATUS"]
    for col in required_lib_cols:
        if col not in df_library.columns:
            st.error(f"Library_data mangler kolonnen {col} til SKU mapping.")
            return None

    # Omdøb library-kolonner
    df_library_renamed = df_library.rename(columns={
        "PRODUCT": "LIB_PRODUCT",
        "EUR ITEM NO.": "LIB_EUR_ITEM_NO",
        "GBP ITEM NO.": "LIB_GBP_ITEM_NO",
        "APMEA ITEM NO.": "LIB_APMEA_ITEM_NO",
        "USD PATTERN NO.": "LIB_USD_PATTERN_NO",
        "MATCH STATUS": "LIB_MATCH_STATUS"
    })

    # Direct merge
    merged_direct = pd.merge(
        df_user,
        df_library_renamed,
        how="left",
        left_on="ARTICLE_NO",
        right_on="LIB_EUR_ITEM_NO"
    )

    # Fallback
    mask_no_direct = merged_direct["LIB_PRODUCT"].isna()
    fallback_rows = merged_direct.loc[mask_no_direct].copy()
    fallback_rows["SHORT_ARTICLE"] = fallback_rows["ARTICLE_NO"].str.split('-').str[0].str.strip().str.upper()

    fallback_merged = pd.merge(
        fallback_rows.drop(columns=["LIB_PRODUCT","LIB_EUR_ITEM_NO","LIB_GBP_ITEM_NO",
                                    "LIB_APMEA_ITEM_NO","LIB_USD_PATTERN_NO","LIB_MATCH_STATUS"]),
        df_library_renamed,
        how="left",
        left_on="SHORT_ARTICLE",
        right_on="LIB_EUR_ITEM_NO"
    )

    for c
