import streamlit as st
import pandas as pd
import re
import difflib

st.set_page_config(page_title="🌍 Country to Region Matcher", layout="wide")
st.title("🌍 Match Region to Country")


region_file = st.file_uploader("📥 Upload Region File (must have 'Country' and 'Region')", type=["xlsx", "csv"])
data_file = st.file_uploader("📥 Upload Data File (must have 'Country')", type=["xlsx", "csv"])


def normalize(text):
    return re.sub(r'\s+', ' ', str(text).strip().lower())


def match_region(country, region_dict):
    if pd.isna(country):
        return ""

    upper_country = str(country).upper()


    if "KINGDOM" in upper_country:
        return "West Europe"
    if "OF AM" in upper_country or "UNITED STATES" in upper_country:
        return "North America"
    if "EMIRATES" in upper_country:
        return "Middle East"

    norm_country = normalize(country)


    matches = difflib.get_close_matches(norm_country, region_dict.keys(), n=1, cutoff=0.8)
    if matches:
        return region_dict[matches[0]]

    # First 5-character match
    for key in region_dict:
        if norm_country[:5] == key[:5]:
            return region_dict[key]

    return ""


if region_file and data_file:

    region_df = pd.read_csv(region_file) if region_file.name.endswith(".csv") else pd.read_excel(region_file)
    region_df.columns = region_df.columns.str.strip()
    region_df = region_df[["Country", "Region"]].dropna()


    region_dict = {normalize(c): r for c, r in zip(region_df["Country"], region_df["Region"])}


    data_df = pd.read_csv(data_file) if data_file.name.endswith(".csv") else pd.read_excel(data_file)
    data_df.columns = data_df.columns.str.strip()

    if "Country" not in data_df.columns:
        st.error("❌ The data file must contain a 'Country' column.")
    else:

        data_df["Region"] = data_df["Country"].apply(lambda x: match_region(x, region_dict))


        cols = data_df.columns.tolist()
        if "Region" in cols and "Country" in cols:
            cols.insert(cols.index("Country") + 1, cols.pop(cols.index("Region")))
            data_df = data_df[cols]

        st.success("✅ Region matching completed!")
        st.dataframe(data_df)

        st.download_button("📥 Download Result", data=data_df.to_csv(index=False), file_name="region_matched.csv", mime="text/csv")
