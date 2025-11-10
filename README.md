# Muuto Product List Generator

Dette projekt er en **Streamlit-applikation** designet til at strukturere, validere og berige pCon produktdata. Appen konverterer pCon-eksportfiler til tre specialiserede outputformater: en præsentationsliste (Word), en ordreimportfil (Excel) og en detaljeret SKU-mapping/masterdata-eksport (Excel).

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
