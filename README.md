# 💡 Muuto Product List Generator

Dette projekt er en **Streamlit-applikation** designet til at hjælpe med at strukturere, validere og berige pCon produktdata. Appen konverterer pCon-eksportfiler til tre specialiserede outputformater: en præsentationsliste (**Word**), en ordreimportfil (**Excel**) og en detaljeret SKU-mapping/masterdata-eksport (**Excel**).

---

## 🚀 Funktioner og Output

Applikationen behandler uploaded pCon-data sammen med to statiske datafiler og genererer følgende:

| Output Rapport | Filnavn | Type | Logik |
| :--- | :--- | :--- | :--- |
| **Præsentationsliste** | `product-list.docx` | Word | Direkte/Fallback Match mod **Library Data** (`PRODUCT`). Sorteres alfabetisk. |
| **Ordreimportfil** | `order-import.xlsx` | Excel (2 kolonner, ingen header) | Indeholder kun `QUANTITY` og det **oprensede** artikelnummer (`BASE_ARTICLE`). |
| **SKU Mapping & Masterdata** | `SKUmapping-masterdata.xlsx` | Excel (2 faner) | Kombinerer brugerdata med **Library Data** (Fane 1) og **Master Data** (Fane 2) ved hjælp af Fallback-logik. |

---

## 🛠️ Opsætning og Kørsel

### 1. Nødvendige Filer (Skal ligge i rodmappen)

Applikationen kræver, at disse filer er tilgængelige lokalt for at kunne udføre opslagene:

* `Library_data.xlsx` (Indeholder SKU-mapping)
* `Muuto_Master_Data_CON_January_2025_EUR.xlsx` (Indeholder Masterdata)

### 2. Python-Biblioteker

Installer de nødvendige afhængigheder:

```bash
pip install streamlit pandas openpyxl python-docx
````

### 3\. Kørsel

Start appen ved hjælp af Streamlit:

https://muuto-pcon-converter.streamlit.app/


-----

## ⚙️ Kernen i Logikken

### A. Dataindlæsning og Preprocessing

Brugerfilen (pCon-eksport) forventes at være enten Excel med fanen **"Article List"** eller CSV. De første 2 rækker springes over.

| Nyt Kolonnenavn | Kilde (0-indeks) | Transformation |
| :--- | :--- | :--- |
| `ARTICLE_NO` | Kolonne **17** | Strippet, Uppercase |
| `QUANTITY` | Kolonne **30** | |
| `SHORT_TEXT` | Kolonne **2** | Strippet, Uppercase |
| `VARIANT_TEXT` | Kolonne **4** | Strippet, Uppercase, NaN til `""` |

### B. Fallback Nøgle (`get_fallback_key`)

Denne funktion renser `ARTICLE_NO` for at muliggøre matchende på baseniveau, når et direkte match mislykkes.

**Logik:**

1.  Deler artikelnummeret ved det første bindestreg (`-`).
2.  Beholder kun den første del (Base-artikelnummer).
3.  Hvis basen starter med `"SPECIAL"`, fjernes dette præfiks.
4.  Returnerer resultatet i **Uppercase**.

| Input (`ARTICLE_NO`) | Output (`BASE_ARTICLE`) |
| :--- | :--- |
| `12345-00-RED` | `12345` |
| `SPECIAL 12345` | `12345` |

-----

## 📄 Detaljeret SKU/Masterdata Output

Filen `SKUmapping-masterdata.xlsx` indeholder to faner, der kombinerer data fra alle tre kilder ved hjælp af **direkte match** efterfulgt af **fallback match**.

### Fane 1: `Item number mapping`

Fokuserer på berigelse fra **Library Data**.

  * **Kolonner fra Brugerdata:** `Quantity in setting`, `Article No.`, `Short Text`, `Variant text`.
  * **Kolonner fra Library Data:** `Product in setting`, `EUR item no.`, `GBP item no.`, `APMEA item no.`, `USD pattern no.`, `Match status`.

### Fane 2: `Master data export`

Fokuserer på at hente alle detaljer fra den centrale **Master Data**-fil.

  * **Kolonner fra Brugerdata:** `Article No.`, `Short Text`, `Variant text`.
  * **Kolonner fra Master Data:** **Alle** kolonner fra den indlæste masterdata-fil.

<!-- end list -->
