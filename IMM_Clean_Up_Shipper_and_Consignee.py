import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="🧠 IMM Machine Model Extractor", layout="wide")
st.title("🏭 IMM Machine Model & Tonnage Extractor")

# === Upload regex pattern file ===
patterns_file = st.file_uploader("📥 Upload regex pattern Excel (regex.xlsx)", type=["xlsx"])
input_file = st.file_uploader("📄 Upload input Excel file with 'Product Description' and 'Supplier'", type=["xlsx"])

if patterns_file and input_file:
    try:
        # Read patterns
        patterns_df = pd.read_excel(patterns_file)
        patterns_df.columns = patterns_df.columns.str.strip()
        machine_patterns = patterns_df['Pattern'].dropna().tolist()

        # Detect header row with 'Product Description'
        header_row = None
        preview = pd.read_excel(input_file, header=None, nrows=30)

        for i, row in preview.iterrows():
            if any(str(cell).strip().lower() == "product description" for cell in row):
                header_row = i
                break

        if header_row is None:
            st.error("❌ Could not find 'Product Description' in the first 30 rows.")
            st.stop()

        df = pd.read_excel(input_file, header=header_row)

        # === Model Extraction Function ===
        def extract_model(description):
            text = str(description).upper().replace("  ", " ")
            for pattern in machine_patterns:
                try:
                    match = re.search(pattern, text)
                except re.error:
                    continue

                if match:
                    groups = match.groups()
                    if not groups:
                        continue
                    series = groups[0].strip().title()
                    parts = [g.strip() for g in groups[1:] if g and g.strip()]
                    if not parts:
                        return series
                    if len(parts) == 1:
                        part = parts[0]
                        if "/" in part:
                            return f"{series} {'-'.join(part.split('/'))}"
                        split_parts = re.findall(r'\d+[A-Z]?|\d{3,5}', part)
                        return f"{series} {'-'.join(split_parts)}" if split_parts else f"{series} {part}"
                    else:
                        cleaned_parts = [re.sub(r"\s+", "", p) for p in parts]
                        return f"{series} {'-'.join(cleaned_parts)}"
            return ""

        # === Tonnage Extraction Function ===
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
                    if len(numbers) >= 2:
                        return round(int(numbers[1]), 1)
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
            except Exception:
                return ""
            return ""

        if "Product Description" not in df.columns:
            st.error("❌ 'Product Description' column not found.")
            st.stop()

        df["Model"] = df["Product Description"].apply(extract_model)
        df["Tonnage"] = df.apply(extract_tonnage, axis=1)

        # Reorder columns
        if "Model" in df.columns:
            model_idx = df.columns.get_loc("Model")
            tonnage = df.pop("Tonnage")
            df.insert(model_idx + 1, "Tonnage", tonnage)

        st.success("✅ Processing complete.")
        st.dataframe(df.head())

        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="⬇️ Download Processed File",
            data=output,
            file_name="machine_model_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
else:
    st.info("📌 Please upload both the regex pattern and the input Excel file to continue.")
