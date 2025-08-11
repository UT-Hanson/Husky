# app.py

import io
import re
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

# -----------------------
# Utilities
# -----------------------
COMMON_SUFFIXES = {
    "inc","inc.","ltd","ltd.","llc","llc.","co","co.","corp","corp.",
    "corporation","company","limited","sa","ag","bv","pty","pte","kg","kgaa",
    "gmbh","plc","srl","oy","ab","aps","sasu","sas","spa","spzoo","sro",
    "bvba","nv","kft","kk","kabushiki","kaisha","pte.","ltd.","pty","ltd"
}

def normalize_name(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s).lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    words = [w for w in s.split() if w not in COMMON_SUFFIXES]
    return " ".join(words).strip()

def first4prefix(name: str) -> str:
    return normalize_name(name).replace(" ", "")[:4]

def pick_sheet_uploader(label: str):
    f = st.file_uploader(label, type=["csv", "xlsx", "xls"])
    df = None
    info = ""
    if f is not None:
        filename = f.name.lower()
        if filename.endswith(".csv"):
            df = pd.read_csv(f)
            info = f"Loaded CSV: **{f.name}**"
        else:
            xf = pd.ExcelFile(f)
            sheet = st.selectbox(f"Select sheet for **{f.name}**", xf.sheet_names, key=f"sheet_{f.name}")
            df = xf.parse(sheet_name=sheet)
            info = f"Loaded Excel: **{f.name}** — Sheet: **{sheet}**"
    return df, info

def build_reference_maps(ref_df: pd.DataFrame, col_company: str, col_website: str):
    ref_df = ref_df.copy()
    ref_df["_norm"] = ref_df[col_company].apply(normalize_name)
    ref_df["_prefix4"] = ref_df[col_company].apply(first4prefix)

    # exact normalized -> website (first non-null)
    exact_map = {}
    for _, r in ref_df.iterrows():
        nm = r["_norm"]
        web = r[col_website]
        if nm and pd.notna(web) and nm not in exact_map:
            exact_map[nm] = str(web)

    # prefix buckets for fallback
    prefix_map = {}
    for _, r in ref_df.iterrows():
        p4 = r["_prefix4"]
        if not p4:
            continue
        web = str(r[col_website]) if pd.notna(r[col_website]) else ""
        prefix_map.setdefault(p4, []).append((r["_norm"], web))

    # for fuzzy
    name_list = ref_df["_norm"].tolist()
    website_by_norm = {}
    for _, r in ref_df.iterrows():
        nm = r["_norm"]
        web = r[col_website]
        if nm and pd.notna(web) and nm not in website_by_norm:
            website_by_norm[nm] = str(web)

    return exact_map, prefix_map, name_list, website_by_norm

def match_fuzzy_then_prefix(target_name: str, exact_map, prefix_map, name_list, website_by_norm, threshold: int):
    nm = normalize_name(target_name)
    if not nm:
        return None

    # 1) Exact normalized
    if nm in exact_map:
        return exact_map[nm]

    # 2) Fuzzy (ordered letters) >= threshold
    best = process.extractOne(nm, name_list, scorer=fuzz.ratio)
    if best:
        best_name, score, _ = best
        if score >= threshold:
            web = website_by_norm.get(best_name)
            if web:
                return web

    # 3) First 4 letters fallback
    p4 = first4prefix(target_name)
    if p4 and p4 in prefix_map:
        candidates = prefix_map[p4]  # list of (norm_name, website)
        # choose best by similarity to break ties
        best_web, best_score = None, -1
        for cnorm, web in candidates:
            score = fuzz.ratio(nm, cnorm)
            if score > best_score and web:
                best_web, best_score = web, score
        if best_web:
            return best_web

    # no match
    return None

def add_company_website(target_df: pd.DataFrame,
                        ref_df: pd.DataFrame,
                        target_company_col: str,
                        ref_company_col: str,
                        ref_website_col: str,
                        threshold: int = 70) -> pd.DataFrame:

    exact_map, prefix_map, name_list, website_by_norm = build_reference_maps(
        ref_df, ref_company_col, ref_website_col
    )

    out = target_df.copy()
    # compute websites
    websites = []
    for _, name in out[target_company_col].items():
        websites.append(
            match_fuzzy_then_prefix(name, exact_map, prefix_map, name_list, website_by_norm, threshold)
        )

    # insert Company Website right after Company Name
    out.insert(out.columns.get_loc(target_company_col) + 1, "Company Website", websites)

    return out

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    with io.BytesIO() as buffer:
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Matched")
        return buffer.getvalue()

# -----------------------
# UI
# -----------------------
st.set_page_config(page_title="Company Website Matcher", page_icon="🔎", layout="wide")
st.title("🔎 Company Website Matcher")

with st.expander("Instructions", expanded=False):
    st.markdown(
        """
        1) Upload a **Reference** file with columns **Company Name** and **Website**  
        2) Upload a **Target** file with at least **Company Name**  
        3) Pick the sheet (if Excel) and confirm columns  
        4) Click **Match & Generate**  
        \n**Matching order:** Exact (normalized) → Fuzzy (≥ threshold) → First 4 letters → else blank.
        """
    )

# File inputs
st.subheader("1) Upload files")
ref_df, ref_info = pick_sheet_uploader("Upload **Reference** file (CSV/XLSX)")
tgt_df, tgt_info = pick_sheet_uploader("Upload **Target** file (CSV/XLSX)")

if ref_df is not None:
    st.caption(ref_info)
    with st.expander("Preview: Reference (first 10 rows)"):
        st.dataframe(ref_df.head(10), use_container_width=True)

if tgt_df is not None:
    st.caption(tgt_info)
    with st.expander("Preview: Target (first 10 rows)"):
        st.dataframe(tgt_df.head(10), use_container_width=True)

# Column selectors
if (ref_df is not None) and (tgt_df is not None):
    st.subheader("2) Map columns")
    cols_ref = list(ref_df.columns)
    cols_tgt = list(tgt_df.columns)

    ref_company_col = st.selectbox("Reference: Company Name column", cols_ref, index=(cols_ref.index("Company Name") if "Company Name" in cols_ref else 0))
    ref_website_col = st.selectbox("Reference: Website column", cols_ref, index=(cols_ref.index("Website") if "Website" in cols_ref else 0))
    tgt_company_col = st.selectbox("Target: Company Name column", cols_tgt, index=(cols_tgt.index("Company Name") if "Company Name" in cols_tgt else 0))

    st.subheader("3) Fuzzy threshold")
    threshold = st.slider("Similarity threshold (used before 1st-4-letters fallback)", 50, 95, 70, 1)

    st.subheader("4) Run")
    if st.button("Match & Generate"):
        with st.spinner("Matching..."):
            result = add_company_website(
                target_df=tgt_df,
                ref_df=ref_df,
                target_company_col=tgt_company_col,
                ref_company_col=ref_company_col,
                ref_website_col=ref_website_col,
                threshold=threshold
            )

        st.success("Done!")
        st.dataframe(result.head(50), use_container_width=True)

        st.download_button(
            label="⬇️ Download as Excel",
            data=to_excel_bytes(result),
            file_name="company_website_matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("Upload both files to continue.")
