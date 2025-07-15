import pandas as pd
import re

# === Load machine patterns from external Excel ===
patterns_file = r'C:\Users\bixiang\Downloads\regex.xlsx'
patterns_df = pd.read_excel(patterns_file)
patterns_df.columns = patterns_df.columns.str.strip()
print("Columns in pattern file:", patterns_df.columns.tolist())

# Load only the 'Pattern' column
machine_patterns = patterns_df['Pattern'].dropna().tolist()

# === Detect correct header row from the input Excel ===
input_file = r'C:\Users\bixiang\Downloads\Book4.xlsx'

# Scan the first 30 rows to find the header row
preview = pd.read_excel(input_file, header=None, nrows=30)

print("\n🔍 Preview of the first 30 rows:")
for i, row in preview.iterrows():
    print(f"Row {i}: {list(row)}")

header_row = None
for i, row in preview.iterrows():
    if any(str(cell).strip().lower() == "product description" for cell in row):
        header_row = i
        break

if header_row is None:
    raise ValueError("❌ Could not find 'Product Description' in the first 30 rows.")

print(f"✅ Detected header row: {header_row}")

# Load the data using the detected header row
df = pd.read_excel(input_file, header=header_row)
print("Columns in input file:", df.columns.tolist())

# === Define the function to extract full machine model from description ===
# === Extract tonnage from model if supplier is NETSTAL ===
# === Define the function to extract model ===
def extract_model(description):
    text = str(description).upper().replace("  ", " ")
    for pattern in machine_patterns:
        try:
            match = re.search(pattern, text)
        except re.error as e:
            print(f"❌ Invalid regex: {pattern}\nError: {e}")
            continue

        if match:
            groups = match.groups()
            if not groups:
                continue

            series = groups[0].strip().title()
            parts = [g.strip() for g in groups[1:] if g and g.strip()]

            if not parts:
                return series  # just return series name if nothing else

            if len(parts) == 1:
                part = parts[0]
                if "/" in part:
                    return f"{series} {'-'.join(part.split('/'))}"
                split_parts = re.findall(r'\d+[A-Z]?|\d{3,5}', part)
                if split_parts:
                    return f"{series} {'-'.join(split_parts)}"
                else:
                    return f"{series} {part}"
            else:
                # handle slash case in first group of multiple
                if "/" in parts[0] and len(parts) == 1:
                    return f"{series} {'-'.join(parts[0].split('/'))}"
                cleaned_parts = [re.sub(r"\s+", "", p) for p in parts]
                return f"{series} {'-'.join(cleaned_parts)}"

    return ""





# === Apply model extraction ===
df["Model"] = df["Product Description"].apply(extract_model)

# === Define tonnage extraction ===
import re

def extract_tonnage(row):
    model = str(row.get("Model", "")).upper()
    supplier = str(row.get("Supplier", "")).upper()

    if not model:
        return ""

    try:
        # NETSTAL: First number before dash
        if "NETSTAL" in supplier:
            m = re.search(r'(\d+)-\d+', model)
            return round(int(m.group(1)) / 10, 1) if m else ""

        # DEMAG / SUMITOMO / UBE: First number
        elif any(s in supplier for s in ["DEMAG", "SUMITOMO", "UBE"]):
            m = re.search(r'(\d{2,5})', model)
            return int(m.group(1)) if m else ""

        # ENGEL: First number divided by 10
        elif "ENGEL" in supplier:
            numbers = re.findall(r'\d{2,5}', model)
            if len(numbers) >= 2:
                return round(int(numbers[1]), 1)

        # ARBURG: Second number divided by 10
        elif "ARBURG" in supplier:
            m = re.findall(r'\d{3,5}', model)
            return round(int(m[1]) / 10, 1) if len(m) >= 2 else ""

        # BMB: First number * 10
        elif "BMB" in supplier:
            m = re.search(r'(\d{2,5})', model)
            return int(m.group(1)) * 10 if m else ""

        # AOKI / NAIGAI: Second number * 10
        elif any(s in supplier for s in ["AOKI", "NAIGAI"]):
            m = re.search(r'AL[-\s]?[\dA-Z]+[-\s]?(\d{2,4})', model)
            return int(m.group(1)) * 10 if m else ""

        # ASB / NISSEI ASB
        elif any(s in supplier for s in ["ASB", "NISSEI"]):
            # Try to extract number after ASB
            m = re.search(r'ASB[-\s]?(\d{2,4})', model)
            if m:
                return int(m.group(1))
            # Fallback: first number in model
            m = re.search(r'(\d{2,4})', model)
            return int(m.group(1)) if m else ""

    except Exception:
        return ""

    return ""



# === Apply and insert Tonnage column ===
df["Tonnage"] = df.apply(extract_tonnage, axis=1)
model_index = df.columns.get_loc("Model")
tonnage_column = df.pop("Tonnage")
df.insert(model_index + 1, "Tonnage", tonnage_column)


# === Apply the function to the 'Product Description' column ===
if "Product Description" not in df.columns:
    raise KeyError("❌ 'Product Description' column not found after loading data.")

df["Model"] = df["Product Description"].apply(extract_model)

# === Save the updated file ===
output_file = r"C:\Users\bixiang\Downloads\Check79.xlsx"
df.to_excel(output_file, index=False)
print(f"✅ Done. File saved as: {output_file}")
