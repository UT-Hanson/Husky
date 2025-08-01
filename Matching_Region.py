import streamlit as st
import pandas as pd
import re
import difflib

st.set_page_config(page_title="🗺️ Buyer Region Mapper", layout="wide")
st.title("🌍 Match Buyer Country to Region")

# === Helper function to load file and select sheet ===
def load_excel_with_sheet_selector(uploaded_file, label):
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = st.selectbox(f"Select sheet from {label}:", xls.sheet_names, key=label)
        df = pd.read_excel(xls, sheet_name=sheet_name)
        return df
    except Exception as e:
        st.error(f"Error loading {label}: {e}")
        return None

# === Normalize helper ===
def normalize(text):
    return re.sub(r'\s+', ' ', str(text).strip().lower())

# === Fuzzy match with priority rules ===
def fuzzy_match_region(buyer_country, region_dict):
    if pd.isna(buyer_country):
        return ""

    norm_buyer = normalize(buyer_country)
    upper_buyer = str(buyer_country).upper()

    # Priority keyword rules
    if "KINGDOM" in upper_buyer:
        return "West Europe"
    if "OF AM" in upper_buyer or "UNITED STATES" in upper_buyer:
        return "North America"
    if "EMIRATES" in upper_buyer:
        return "Middle East"

    # Fuzzy match
    matches = difflib.get_close_matches(norm_buyer, region_dict.keys(), n=1, cutoff=0.8)
    if matches:
        return region_dict[matches[0]]

    # Match first 5 characters
    for key in region_dict:
        if norm_buyer[:5] == key[:5]:
            return region_dict[key]

    return ""

# === Upload files ===
base_file = st.file_uploader("📄 Upload Base File (must include 'Buyer Country')", type=["xlsx"])
region_file = st.file_uploader("🌍 Upload Region Match File (with 'Country' and 'Region')", type=["xlsx"])

if base_file and region_file:
    df = load_excel_with_sheet_selector(base_file, "Base File")
    if df is not None and "Buyer Country" in df.columns:
        region_df = load_excel_with_sheet_selector(region_file, "Region File")
        if region_df is not None:
            region_df.columns = region_df.columns.str.strip()
            region_df = region_df[['Country', 'Region']].dropna()

            # Create lookup dictionary
            region_dict = {normalize(country): region for country, region in zip(region_df['Country'], region_df['Region'])}

            # Generate Buyer Region column
            df['Buyer Region'] = df['Buyer Country'].apply(lambda x: fuzzy_match_region(x, region_dict))

            # Reorder Buyer Region next to Buyer Country
            # Reorder Buyer Region next to Buyer Country
            # Reorder Buyer Region next to Buyer Country
            country_index = df.columns.get_loc("Buyer Country")
            cols = df.columns.tolist()
            cols.insert(country_index + 1, cols.pop(cols.index("Buyer Region")))
            df = df[cols]

            st.success("✅ Region mapping complete!")
            st.dataframe(df)

            # --- Fix for export ---
            from io import BytesIO

            output = BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.download_button(
                label="📥 Download Updated File",
                data=output,
                file_name="buyer_with_region.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


        else:
            st.error("⚠️ Region file must contain 'Country' and 'Region' columns.")
    else:
        st.error("⚠️ Base file must contain 'Buyer Country' column.")
