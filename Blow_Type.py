# app.py
# Streamlit: Classify Product Type Focus (Single Stage / Linear Blower)
# Rules:
# - Supplier Cleaned Final in {1BLOW SAS, CHUMPOWER MACHINERY, SIAPI, SIDE INDIA} => Linear Blower
# - Supplier Cleaned Final in {AOKI TECHNICAL LABORATORY INC, NISSEI ASB} => Single Stage
# - Supplier Cleaned Final == SIPA:
#       If Product Description contains 'Linear' or 'XTRA' => Linear Blower
#       If Product Description contains 'ECS'            => Single Stage
# - Otherwise blank
# Output column is inserted immediately after "Product Description".

import io
import pandas as pd
import streamlit as st

LINEAR_BLOWER_SUPPLIER_KEYWORDS = {
    "1BLOW", "CHUMPOWER", "SIAPI", "SIDE INDIA"  # add more keywords if needed
}

SINGLE_STAGE_SUPPLIER_KEYWORDS = {
    "AOKI", "NISSEI", "ASB"   # as requested
}

NEW_COL = "Product Type Focus (Single Stage/ Linear Blower)"

st.set_page_config(page_title="Product Type Classifier", page_icon="🧪", layout="wide")
st.title("🧪 Product Type Classifier")

st.markdown(
    "Upload a XLSX with **Supplier Cleaned Final** and **Product Description**. "
    "The app adds **Product Type Focus (Single Stage/ Linear Blower)** based on your rules."
)

def pick_sheet_uploader(label: str):
    f = st.file_uploader(label, type=["csv", "xlsx", "xls"])
    df = None
    info = ""
    if f is not None:
        name = f.name.lower()
        if name.endswith(".csv"):
            df = pd.read_csv(f)
            info = f"Loaded CSV: **{f.name}**"
        else:
            xf = pd.ExcelFile(f)
            sheet = st.selectbox(f"Select sheet for **{f.name}**", xf.sheet_names, key=f"sheet_{f.name}")
            df = xf.parse(sheet_name=sheet)
            info = f"Loaded Excel: **{f.name}** — Sheet: **{sheet}**"
    return df, info

def classify_row(supplier, description) -> str:
    # safe normalize
    s = str(supplier).upper() if pd.notna(supplier) else ""
    d = str(description).upper() if pd.notna(description) else ""

    # SIPA special rule first
    if "SIPA" in s:
        if ("LINEA" in d) or ("XTRA" in d) or ("SFL" in d):
            return "Linear Blower"
        if "ECS" in d:
            return "Single Stage"
        # if SIPA but no keyword, fall through to generic contains checks

    # Contains-match for Single Stage (AOKI/NISSEI/ASB)
    if any(k in s for k in SINGLE_STAGE_SUPPLIER_KEYWORDS):
        return "Single Stage"

    # Contains-match for Linear Blower
    if any(k in s for k in LINEAR_BLOWER_SUPPLIER_KEYWORDS):
        return "Linear Blower"

    # No decision
    return ""


def insert_next_to(df: pd.DataFrame, after_col: str, new_col: str, values) -> pd.DataFrame:
    out = df.copy()
    pos = out.columns.get_loc(after_col) + 1
    out.insert(pos, new_col, values)
    return out

df, info = pick_sheet_uploader("Upload file (CSV/XLSX)")
if df is not None:
    st.caption(info)
    missing = [c for c in ["Supplier Cleaned Final", "Product Description"] if c not in df.columns]
    if missing:
        st.error(f"Missing required column(s): {', '.join(missing)}")
    else:
        with st.expander("Preview (first 20 rows)"):
            st.dataframe(df.head(20), use_container_width=True)

        if st.button("Classify"):
            with st.spinner("Classifying..."):
                vals = [classify_row(r["Supplier Cleaned Final"], r["Product Description"]) for _, r in df.iterrows()]
                result = insert_next_to(df, "Product Description", NEW_COL, vals)

            st.success("Done! Column added next to Product Description.")
            st.dataframe(result.head(50), use_container_width=True)

            # Download as Excel
            with io.BytesIO() as buf:
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    result.to_excel(writer, index=False, sheet_name="Classified")
                data = buf.getvalue()

            st.download_button(
                "⬇️ Download Excel",
                data=data,
                file_name="product_type_classified.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
else:
    st.info("Upload a file to begin.")

