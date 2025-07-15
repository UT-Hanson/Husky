import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
from rapidfuzz import fuzz
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.set_page_config(page_title="🧼 CRM Account Cleanup", layout="wide")
st.title("🔍 CRM Account Cleanup & Deduplication Tool")

# === Cleaning Functions ===
def clean_name(name):
    if pd.isna(name):
        return ""
    name = unicodedata.normalize("NFKD", str(name))
    name = name.encode("ascii", "ignore").decode("utf-8")
    name = name.lower()
    name = re.sub(r'[^\w\s]', '', name)
    name = re.sub(r'\b(inc|llc|ltd|co|corporation|company|limited|group|plc|gmbh|sa|bv|global|sarl|sro|kg|ltda|s de rl|operations|applied to life|automotive)\b', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def classify_sou(sou):
    if pd.isna(sou): return "OTHER"
    sou = sou.lower()
    if "hot runners" in sou or "hrc" in sou:
        return "HOT RUNNERS"
    elif "beverage" in sou or "packaging" in sou:
        return "PACKAGING"
    elif "msp" in sou or "csm" in sou:
        return "OTHER PACKAGING"
    return "OTHER"

def detect_header_row(file, sheet_name, required_cols):
    for i in range(10):
        try:
            df = pd.read_excel(file, sheet_name=sheet_name, header=i, nrows=1)
            if all(col in df.columns for col in required_cols):
                return i
        except:
            continue
    return None

# === File Upload ===
uploaded_file = st.file_uploader("📂 Upload CRM Excel File", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("📑 Select sheet:", xls.sheet_names)

    required_columns = ['Account Name', 'SOU', 'Country', 'Business Type']
    header_row = detect_header_row(uploaded_file, sheet_name, required_columns)

    if header_row is None:
        st.error("❌ Could not detect required columns in the first 10 rows.")
        st.stop()

    full_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
    total_rows = len(full_df)

    # Select row range
    st.subheader("📊 Row Range Selection")
    start_row = st.number_input("Start row (0-based):", min_value=0, max_value=total_rows - 1, value=0)
    end_row = st.number_input("End row (exclusive):", min_value=start_row + 1, max_value=total_rows, value=min(start_row + 5000, total_rows))

    if st.button("🔄 Run Cleanup"):
        df = full_df.iloc[start_row:end_row].reset_index(drop=True)

        # Step 1: Clean names
        df['Cleaned Name'] = df['Account Name'].apply(clean_name)

        # Step 2: SOU category
        df['SOU Category'] = df['SOU'].apply(classify_sou)

        # Step 3: Group names by fuzzy match
        group_names = [''] * len(df)
        assigned = set()

        for i in range(len(df)):
            if i in assigned:
                continue
            base_name = df.loc[i, 'Cleaned Name']
            base_sou = df.loc[i, 'SOU Category']
            base_prefix = " ".join(base_name.split()[:2])
            group_names[i] = base_prefix
            assigned.add(i)

            for j in range(i + 1, len(df)):
                if j in assigned:
                    continue
                if df.loc[j, 'SOU Category'] != base_sou:
                    continue
                score = fuzz.token_set_ratio(base_name, df.loc[j, 'Cleaned Name'])
                if score > 85:
                    group_names[j] = base_prefix
                    assigned.add(j)

        df['Account Group Name Cleaned'] = group_names

        # Step 4: Deduplicate by Cleaned Name + Country
        deduped = []
        skip = set()

        for i in range(len(df)):
            if i in skip:
                continue
            current = df.iloc[i]
            current_name = current['Cleaned Name']
            current_country = current['Country']
            similar = [i]

            for j in range(i + 1, len(df)):
                if df.iloc[j]['Country'] != current_country:
                    continue
                score = fuzz.token_set_ratio(current_name, df.iloc[j]['Cleaned Name'])
                if score > 90:
                    similar.append(j)
                    skip.add(j)

            group = df.iloc[similar]
            customers = group[group['Business Type'] == 'Customer']
            best = customers.iloc[0] if not customers.empty else group.iloc[0]
            deduped.append(best)

        final_df = pd.DataFrame(deduped)

        # Step 5: Reorder columns
        preferred_start = [
            'Account Name',
            'Account Group Name Cleaned',
            'Business Type',
            'Country',
            'SOU Category',
            'SOU',
            'Owner'
        ]
        front = [col for col in preferred_start if col in final_df.columns]
        others = [col for col in final_df.columns if col not in front]
        final_df = final_df[front + others]

        # Step 6: Drop unwanted system columns
        cols_to_drop = ['(Do Not Modify) Account', '(Do Not Modify) Row Checksum']
        final_df = final_df.drop(columns=[col for col in cols_to_drop if col in final_df.columns])

        # Step 7: Export cleaned Excel
        output = BytesIO()
        final_df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.success(f"✅ Cleanup complete! Processed {len(df)} rows → deduplicated to {len(final_df)} rows.")
        st.download_button(
            label="⬇️ Download Cleaned Excel",
            data=output.getvalue(),
            file_name="cleaned_accounts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
