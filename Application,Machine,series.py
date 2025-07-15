import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="🛠️ IMM Matcher", layout="wide")
st.title("🔧 IMM Model, Tonnage, and Application Matcher")

# === Upload Files ===
patterns_file = st.file_uploader("📥 Upload regex pattern file", type=["xlsx"])
application_file = st.file_uploader("📥 Upload application keyword file", type=["xlsx"])
input_file = st.file_uploader("📥 Upload input Excel file", type=["xlsx"])

if patterns_file and application_file and input_file:
    try:
        # === Load Regex Patterns ===
        patterns_df = pd.read_excel(patterns_file)
        patterns_df.columns = patterns_df.columns.str.strip()
        machine_patterns = patterns_df['Pattern'].dropna().tolist()

        # === Load Application Mapping ===
        app_map_df = pd.read_excel(application_file)
        app_map_df.columns = app_map_df.columns.str.strip()
        application_keywords = dict(zip(app_map_df['Keyword'].str.upper(), app_map_df['Application']))

        # === Detect Header Row in Main Input File ===
        preview = pd.read_excel(input_file, header=None, nrows=30)
        header_row = None
        for i, row in preview.iterrows():
            if any(str(cell).strip().lower() == "product description" for cell in row):
                header_row = i
                break
        if header_row is None:
            st.error("❌ Could not find 'Product Description' column.")
            st.stop()
        df = pd.read_excel(input_file, header=header_row)

        # === Extract Model ===
        def extract_model(description):
            text = str(description).upper().replace("  ", " ")
            for pattern in machine_patterns:
                try:
                    match = re.search(pattern, text)
                    if not match:
                        continue
                    groups = match.groups()
                    if not groups:
                        continue
                    series = groups[0].strip().title()
                    parts = [g.strip() for g in groups[1:] if g and g.strip()]
                    if not parts:
                        return series
                    if len(parts) == 1:
                        if "/" in parts[0]:
                            return f"{series} {'-'.join(parts[0].split('/'))}"
                        split_parts = re.findall(r'\d+[A-Z]?|\d{3,5}', parts[0])
                        return f"{series} {'-'.join(split_parts)}" if split_parts else f"{series} {parts[0]}"
                    cleaned_parts = [re.sub(r"\s+", "", p) for p in parts]
                    return f"{series} {'-'.join(cleaned_parts)}"
                except re.error:
                    continue
            return ""

        df["Model"] = df["Product Description"].apply(extract_model)

        # === Extract Tonnage ===
        def extract_tonnage(row):
            model = str(row.get("Model", "")).upper()
            supplier = str(row.get("Supplier", "")).upper()
            if not model:
                return ""
            try:
                if "NETSTAL" in supplier:
                    m = re.search(r'(\d+)-\d+', model)
                    return round(int(m.group(1)) / 10, 1) if m else ""
                elif any(s in supplier for s in ["DEMAG", "SUMITOMO", "UBE"]):
                    m = re.search(r'(\d{2,5})', model)
                    return int(m.group(1)) if m else ""
                elif "ENGEL" in supplier:
                    numbers = re.findall(r'\d{2,5}', model)
                    return round(int(numbers[1]), 1) if len(numbers) >= 2 else ""
                elif "ARBURG" in supplier:
                    m = re.findall(r'\d{3,5}', model)
                    return round(int(m[1]) / 10, 1) if len(m) >= 2 else ""
                elif "BMB" in supplier:
                    m = re.search(r'(\d{2,5})', model)
                    return int(m.group(1)) * 10 if m else ""
                elif any(s in supplier for s in ["AOKI", "NAIGAI"]):
                    m = re.search(r'AL[-\s]?[\dA-Z]+[-\s]?(\d{2,4})', model)
                    return int(m.group(1)) * 10 if m else ""
                elif any(s in supplier for s in ["ASB", "NISSEI"]):
                    m = re.search(r'ASB[-\s]?(\d{2,4})', model)
                    if m:
                        return int(m.group(1))
                    m = re.search(r'(\d{2,4})', model)
                    return int(m.group(1)) if m else ""
            except:
                return ""
            return ""

        df["Tonnage"] = df.apply(extract_tonnage, axis=1)

        # === Match Application ===
        def match_application(model):
            text = str(model).upper()
            text = re.sub(r"[^A-Z0-9]", "", text)
            matched_keywords = []
            for keyword, application in application_keywords.items():
                normalized_key = re.sub(r"[^A-Z0-9]", "", keyword.upper())
                if normalized_key in text:
                    matched_keywords.extend(application.split(";"))
            unique_apps = sorted(set(matched_keywords))
            if "Other" in unique_apps and len(unique_apps) > 1:
                unique_apps.remove("Other")
            return ";".join(unique_apps)

        df["Application"] = df["Model"].apply(match_application)

        # === Reorder Columns ===
        model_index = df.columns.get_loc("Model")
        df.insert(model_index + 1, "Tonnage", df.pop("Tonnage"))
        df.insert(model_index + 2, "Application", df.pop("Application"))

        # === Show and Download Output ===
        st.success("✅ Matching Complete!")
        st.dataframe(df.head())

        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="⬇️ Download Processed Excel",
            data=output,
            file_name="IMM_Matched_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("📌 Please upload all three files to begin.")
