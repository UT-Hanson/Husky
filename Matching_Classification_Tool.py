import streamlit as st
import pandas as pd
import re
from io import BytesIO
from rapidfuzz import fuzz

st.set_page_config(page_title="CRM Matching Tool", layout="wide")
st.title("🔍 CRM Matching & Classification Tool")

# === Upload files ===
database_file = st.file_uploader("📂 Upload CRM Database File (Excel)", type=["xlsx"])
lookup_file = st.file_uploader("📂 Upload Lookup File (Excel)", type=["xlsx"])

def detect_header_row(file, required_cols, file_label):
    """Detect the row that contains all required columns."""
    xl = pd.ExcelFile(file)
    sheet_names = xl.sheet_names
    sheet_name = st.selectbox(f"Select sheet from {file_label}:", sheet_names, key=file.name)

    for i in range(10):  # Check first 10 rows
        try:
            df_sample = pd.read_excel(file, sheet_name=sheet_name, header=i, nrows=1)
            if all(col in df_sample.columns for col in required_cols):
                st.success(f"✅ Header detected at row {i + 1} in {file_label}")
                return sheet_name, i
        except Exception:
            continue
    return None, None

if database_file and lookup_file:
    required_db_columns = ['Account Name', 'Account Group Name Cleaned', 'Country', 'Business Type']

    required_lookup_columns = ['Company name', 'Country']

    st.subheader("📋 CRM File Settings")
    db_sheet, db_header = detect_header_row(database_file, required_db_columns, "CRM File")

    st.subheader("📋 Lookup File Settings")
    lookup_sheet, lookup_header = detect_header_row(lookup_file, required_lookup_columns, "Lookup File")

    if db_sheet is None:
        st.error("❌ Could not find required columns in CRM file.")
        st.stop()
    if lookup_sheet is None:
        st.error("❌ Could not find required columns in Lookup file.")
        st.stop()

    # Load data with detected headers
    db = pd.read_excel(database_file, sheet_name=db_sheet, header=db_header)
    lookup = pd.read_excel(lookup_file, sheet_name=lookup_sheet, header=lookup_header)

    # === Cleaning functions ===
    def clean_name(name):
        if pd.isna(name): return ''
        name = name.lower()
        name = re.sub(r'[^\w\s]', '', name)
        name = re.sub(r'\b(inc|llc|ltd|co|corporation|company|limited|group|division|plc|gmbh|sa|bv|global|sarl|sro|kg|ltda|operations|applied to life|automotive|packaging)\b', '', name)
        name = re.sub(r'\s+', ' ', name).strip()
        return name

    def get_prefix(name):
        return re.sub(r'\s+', '', name)[:4]


    db['Cleaned Account Name'] = db['Account Name'].apply(clean_name)
    db['Cleaned Group Name'] = db['Account Group Name Cleaned'].apply(clean_name)
    db['Cleaned Name'] = db['Cleaned Account Name']  # Use Account Name as backup for Cleaned Name
    db['Prefix'] = db['Cleaned Account Name'].apply(get_prefix)

    lookup['Cleaned Company Name'] = lookup['Company name'].apply(clean_name)
    lookup['Prefix'] = lookup['Cleaned Company Name'].apply(get_prefix)

    # === Matching logic ===
    def get_threshold(name_len):
        return 88 if name_len <= 12 else 80 if name_len <= 20 else 70

    def match_business_type(cleaned_name, country, prefix):
        if prefix not in set(db['Prefix']):
            return "Not in CRM"

        best_score = 0
        best_type = None
        same_country_type = None
        other_region_customer = False
        threshold = get_threshold(len(cleaned_name))

        for _, row in db.iterrows():
            if prefix != row['Prefix']:
                continue

            db_country = str(row['Country']).strip().lower()
            db_type = str(row['Business Type']).strip()
            scores = [
                fuzz.token_set_ratio(cleaned_name, row['Cleaned Account Name']),
                fuzz.token_set_ratio(cleaned_name, row['Cleaned Group Name']),
                fuzz.token_set_ratio(cleaned_name, row['Cleaned Name'])
            ]
            score = max(scores)

            if score >= 90 and db_country == country:
                return db_type
            if score >= 90 and db_country != country and db_type == "Customer":
                other_region_customer = True
            if score >= threshold and db_country == country and same_country_type is None:
                same_country_type = db_type
            if score > best_score:
                best_score = score
                best_type = db_type

        if same_country_type:
            return same_country_type
        if other_region_customer:
            return "Not in CRM (Customer in other region)"
        if best_score >= (threshold - 5):
            return "Other"
        return "Not in CRM"

    st.write("🔄 Matching business types...")
    lookup['Business Type Matched'] = lookup.apply(
        lambda row: match_business_type(
            row['Cleaned Company Name'],
            str(row['Country']).strip().lower(),
            row['Prefix']
        ), axis=1
    )

    year_columns = sorted([col for col in lookup.columns if str(col).isdigit()])
    num_years = len(year_columns)
    last_year = year_columns[-1]


    def classify_company(row):
        business_type = str(row.get("Business Type Matched", "")).strip().lower()
        company_name = str(row.get("Company name", "")).strip().lower()


        values = [row.get(y, 0) if pd.notna(row.get(y)) else 0 for y in year_columns]
        grand_total = sum(values)

        last_year_value = row.get(last_year)
        last_year_value = 0 if pd.isna(last_year_value) else last_year_value


        if any(kw in business_type for kw in ["customer", "logistics", "trading"]) or \
                any(kw in company_name for kw in ["logistics", "freight", "trading", "forwarding", "export"]):
            return "D"


        if "prospect" in business_type:
            return "P"


        if last_year_value > 20000:
            previous_years_values = values[:-1]
            if all(v == 0 for v in previous_years_values) or \
                    (max(previous_years_values) < 5000 and last_year_value > 30000):
                return "B"


        if last_year_value > 0:
            non_zero_values = [v for v in values if v > 0]
            if len(non_zero_values) >= 2:
                increasing = all(non_zero_values[i] >= non_zero_values[i - 1] for i in range(1, len(non_zero_values)))
                if increasing:
                    if grand_total >= num_years * 10000 or \
                            (last_year_value > 20000 and last_year_value >= 0.6 * grand_total):
                        return "C"


        if grand_total >= num_years * 30000 and last_year_value > 30000:
            return "A"


        import_years = sum(1 for v in values if v > 0)
        latest_year_imported = last_year_value > 0

        if (grand_total >= num_years * 30000 and last_year_value <= 30000) or \
                (latest_year_imported and import_years >= 0.8 * num_years and grand_total >= 20000 * num_years):
            return "F"


        return "N"


    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter

    st.write("🔠 Classifying companies...")
    lookup['Classification'] = lookup.apply(classify_company, axis=1)


    lookup.drop(columns=["Cleaned Company Name", "Prefix"], inplace=True, errors="ignore")

    st.success("✅ Finished processing!")


    temp_output = BytesIO()
    lookup.to_excel(temp_output, index=False, engine='openpyxl')
    temp_output.seek(0)


    wb = load_workbook(temp_output)
    ws = wb.active

    description = (
        "Classification:\n"
        "D = Customer/Logistics/Trading\n"
        "P = Prospect in CRM\n"
        "B = New and sudden growth >20K in latest year\n"
        "C = Increasing trend +  (Total >10K/year or (60% in latest year > 20k))\n"
        "A = High-performing (Total >30K/year and latest year >30K)\n"
        "F = Stable or reduced recently (Total > 30k/ year and latest year <30k) or (≥80% active years including latest year, and total >20K/year)\n"
        "N = Not interesting to focus (Latest <20k and total <30k/year and no increasing trend and <80% including latest year)"
    )

    ws.insert_rows(1)
    end_col = min(10, ws.max_column)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)

    desc_cell = ws.cell(row=1, column=1)
    desc_cell.value = description
    desc_cell.fill = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")  # Light yellow
    desc_cell.font = Font(bold=True)


    header_row = 2
    highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        col_value = ws.cell(row=header_row, column=col).value
        if col_value in ["Business Type Matched", "Classification"]:
            ws.cell(row=header_row, column=col).fill = highlight


    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    st.download_button(
        label="⬇️ Download Final Excel",
        data=final_output.getvalue(),
        file_name="Processed_Lookup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
