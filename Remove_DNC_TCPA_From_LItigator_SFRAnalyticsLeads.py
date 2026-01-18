import os
import re
import pandas as pd

# ======================================
# Folder paths
# ======================================
base_dir = r"C:\Users\awuzw\Desktop\Python Scripts\SFRAnalytics_TCPA_Cleanup"

dnc_folder = os.path.join(base_dir, "TCPA - DNC Matches")
primary_folder = os.path.join(base_dir, "TCPA - SFRAnalytics Primary")
output_folder = os.path.join(base_dir, "TCPA - Output")

for folder in (dnc_folder, primary_folder, output_folder):
    os.makedirs(folder, exist_ok=True)

# ======================================
# Helper: Normalize phone → 10 digits
# ======================================
def normalize_phone(value):
    if pd.isna(value):
        return ''

    if isinstance(value, (int, float)):
        try:
            s = str(int(value))
        except:
            s = str(value)
    else:
        s = str(value)

    if re.match(r'^\d+\.0$', s):
        s = s[:-2]

    digits = re.sub(r'\D', '', s)

    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    return digits if len(digits) == 10 else ''

# ======================================
# Load DNC Excel file
# ======================================
dnc_files = [f for f in os.listdir(dnc_folder)
             if f.lower().endswith(('.xlsx', '.xls'))]

print("\n=== FILES FOUND IN TCPA - DNC MATCHES FOLDER ===")
for f in os.listdir(dnc_folder):
    print(" -", f)
print(len(dnc_files))


if len(dnc_files) != 1:
    raise ValueError(f"Expected exactly ONE Excel file in {dnc_folder}")

dnc_path = os.path.join(dnc_folder, dnc_files[0])
print(f"\n=== USING DNC MATCH FILE ===\n{dnc_path}\n")

df_dnc = pd.read_excel(dnc_path)

print("=== DNC FILE HEADERS ===")
for i, c in enumerate(df_dnc.columns):
    print(f"{i:02d}: {c}")
print()

# ======================================
# Build DNC/TCPA phone set (verbose)
# ======================================
high_risk_numbers = set()

phone_columns = ["First Phone", "Second Phone", "Third Phone"]

print("=== BEGIN ROW-BY-ROW DNC SCAN ===\n")

for idx, row in df_dnc.iterrows():

    borrower = row.get("Borrower", "")
    owner = row.get("Owner Name", "")

    print(f"\nROW {idx}")
    print(f"Borrower   : {borrower}")
    print(f"Owner Name : {owner}")

    for phone_col in phone_columns:
        if phone_col not in df_dnc.columns:
            continue

        phone_raw = row[phone_col]
        phone_norm = normalize_phone(phone_raw)

        # Identify right-hand cell
        col_idx = df_dnc.columns.get_loc(phone_col)
        right_val = None
        right_col = None

        if col_idx + 1 < len(df_dnc.columns):
            right_col = df_dnc.columns[col_idx + 1]
            right_val = row[right_col]

        if not phone_norm:
            print(f"  {phone_col}: [NO VALID PHONE]")
            continue

        print(f"  {phone_col}: {phone_norm}")

        flag = ""
        if right_val is not None and not pd.isna(right_val):
            tag = str(right_val).lower()
            if "tcpa" in tag:
                flag = "TCPA ⚠️⚠️⚠️"
            elif "dnc" in tag:
                flag = "DNC ⚠️"

        if flag:
            print(f"    → RIGHT CELL [{right_col}] = {right_val}  ==> FLAGGED: {flag}")
            high_risk_numbers.add(phone_norm)
        else:
            print(f"    → RIGHT CELL [{right_col}] = {right_val}")

print("\n=== DNC SCAN COMPLETE ===")
print(f"Total HIGH-RISK phone numbers captured: {len(high_risk_numbers)}")
print("Sample:", list(high_risk_numbers)[:20])
