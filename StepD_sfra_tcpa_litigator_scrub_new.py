import os
import re
import pandas as pd

# ======================================
# Folder paths
# ======================================
base_dir = r"C:\Users\awuzw\Desktop\Python Scripts\SFRAnalytics_TCPA_Cleanup"

dnc_folder = os.path.join(base_dir, "StepD - Input TCPA Liti DNC Match")
primary_folder = os.path.join(base_dir, "StepD - Input SFRA Leads")
output_folder = os.path.join(base_dir, "StepD - Output SFRA Leads Scrubbed")

for folder in (dnc_folder, primary_folder, output_folder):
    os.makedirs(folder, exist_ok=True)

# ======================================
# Helper: Normalize phone → 10 digits
# ======================================
def normalize_phone(value):
    if pd.isna(value):
        return ""

    s = str(value).strip()

    # Masked TCPA phone like ******1655
    if re.match(r"^\*+\d{2,4}$", s):
        return ""

    if isinstance(value, (int, float)):
        try:
            s = str(int(value))
        except:
            pass

    if re.match(r"^\d+\.0$", s):
        s = s[:-2]

    digits = re.sub(r"\D", "", s)

    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    return digits if len(digits) == 10 else ""

# ======================================
# Load DNC / TCPA Matches Excel
# ======================================
dnc_files = [
    f for f in os.listdir(dnc_folder)
    if f.lower().endswith((".xlsx", ".xls"))
]

print("\n=== FILES FOUND IN TCPA - DNC MATCHES FOLDER ===")
for f in os.listdir(dnc_folder):
    print(" -", f)
print("Excel files found:", len(dnc_files))

if len(dnc_files) != 1:
    raise ValueError("Expected exactly ONE DNC Excel file")

dnc_path = os.path.join(dnc_folder, dnc_files[0])
print(f"\n=== USING DNC MATCH FILE ===\n{dnc_path}\n")

df_dnc = pd.read_excel(dnc_path)

print("=== DNC FILE HEADERS ===")
for i, c in enumerate(df_dnc.columns):
    print(f"{i:02d}: {c}")
print()

# ======================================
# Build DNC + TCPA sets (VERBOSE)
# ======================================
dnc_numbers = set()
tcpa_numbers = set()
tcpa_borrowers = set()
tcpa_owners = set()

phone_columns = [
    "First Phone",
    "Second Phone",
    "Third Phone",
    "Propstream Phone 1",
    "Propstream Phone 2",
    "Propstream Phone 3",
    "Propstream Phone 4",
    "Propstream Phone 5"
]

print("=== BEGIN ROW-BY-ROW DNC / TCPA SCAN ===\n")

for idx, row in df_dnc.iterrows():

    borrower = str(row.get("Borrower", "")).strip()
    owner = str(row.get("Owner Name", "")).strip()

    print(f"\nROW {idx}")
    print(f"Borrower   : {borrower}")
    print(f"Owner Name : {owner}")

    for phone_col in phone_columns:
        if phone_col not in df_dnc.columns:
            continue

        raw_phone = str(row.get(phone_col, "")).strip()
        phone_norm = normalize_phone(raw_phone)

        col_idx = df_dnc.columns.get_loc(phone_col)
        right_col = None
        right_val = None

        if col_idx + 1 < len(df_dnc.columns):
            right_col = df_dnc.columns[col_idx + 1]
            right_val = row.iloc[col_idx + 1]

        if not raw_phone:
            print(f"  {phone_col}: [EMPTY]")
            continue

        print(f"  {phone_col}: {raw_phone}")

        if right_val is not None and not pd.isna(right_val):
            tag = str(right_val).lower()

            # if "tcpa" in tag:
            if "tcpa" in tag or "dnc_complainers" in tag:
                print(f"    → RIGHT CELL [{right_col}] = {right_val} ==> TCPA ⚠️⚠️⚠️")

                tcpa_borrowers.add(borrower)
                tcpa_owners.add(owner)

                if phone_norm:
                    tcpa_numbers.add(phone_norm)

            elif "dnc" in tag:
                print(f"    → RIGHT CELL [{right_col}] = {right_val} ==> DNC ⚠️")

                if phone_norm:
                    dnc_numbers.add(phone_norm)
            else:
                print(f"    → RIGHT CELL [{right_col}] = {right_val}")
        else:
            print(f"    → RIGHT CELL [{right_col}] = {right_val}")

print("\n=== DNC / TCPA SET SUMMARY ===")
print(f"DNC phone numbers      : {len(dnc_numbers)}")
print(f"TCPA phone numbers     : {len(tcpa_numbers)}")
print(f"TCPA borrowers blocked : {len(tcpa_borrowers)}")
print(f"TCPA owners blocked    : {len(tcpa_owners)}")

# ======================================
# Load SFRAnalytics Primary CSV
# ======================================
primary_files = [
    f for f in os.listdir(primary_folder)
    if f.lower().endswith(".csv")
]

if len(primary_files) != 1:
    raise ValueError("Expected exactly ONE primary CSV")

primary_path = os.path.join(primary_folder, primary_files[0])
print(f"\n=== USING SFRANALYTICS PRIMARY FILE ===\n{primary_path}\n")

df_primary = pd.read_csv(primary_path, dtype=str)

# ======================================
# Scrub primary with HARD TCPA ROW DROP
# ======================================
kept_rows = []
rows_dropped_tcpa = 0

original_phone_cells = 0
kept_phone_cells = 0
phones_removed = 0

for _, row in df_primary.iterrows():

    borrower = str(row.get("Borrower LLC", "")).strip()
    owner = str(row.get("Owner Name", "")).strip()

    print(borrower)

    # Count ORIGINAL phone cells (before any exclusion)
    for col in phone_columns:
        if col in row and normalize_phone(row.get(col)):
            original_phone_cells += 1

    # HARD TCPA EXCLUSION
    if borrower in tcpa_borrowers or owner in tcpa_owners:
        rows_dropped_tcpa += 1
        continue

    row = row.copy()

    # Phone-level scrubbing
    for col in phone_columns:
        if col not in row:
            continue

        norm_val = normalize_phone(row.get(col))
        if not norm_val:
            continue

        if norm_val in dnc_numbers or norm_val in tcpa_numbers:
            row[col] = ""
            phones_removed += 1

    # Count KEPT phone cells
    for col in phone_columns:
        if normalize_phone(row.get(col)):
            kept_phone_cells += 1

    kept_rows.append(row)

df_clean = pd.DataFrame(kept_rows)

# ======================================
# Lead stats
# ======================================
leads_with_phone = 0
leads_without_phone = 0

for _, row in df_clean.iterrows():
    if any(normalize_phone(row.get(col)) for col in phone_columns):
        leads_with_phone += 1
    else:
        leads_without_phone += 1

# ======================================
# Save output
# ======================================
base, ext = os.path.splitext(primary_files[0])
out_path = os.path.join(output_folder, base + "_phones_scrubbed" + ext)
df_clean.to_csv(out_path, index=False)

# ======================================
# TCPA MANUAL REVIEW PRINT
# ======================================
print("\n=== TCPA MANUAL REVIEW LIST ===")

print("\nTCPA Borrower Entities (HARD BLOCKED):")
if tcpa_borrowers:
    for b in sorted(tcpa_borrowers):
        print(" -", b)
else:
    print(" (none)")

print("\nTCPA Owner Names (HARD BLOCKED):")
if tcpa_owners:
    for o in sorted(tcpa_owners):
        print(" -", o)
else:
    print(" (none)")

# ======================================
# FINAL REPORT
# ======================================
print("\n=== TCPA HARD-EXCLUDED SUMMARY ===")
print(f"Rows dropped entirely due to TCPA: {rows_dropped_tcpa}")

print("\n=== FINAL STATS ===")
print(f"Original rows            : {len(df_primary)}")
print(f"Rows kept                : {len(df_clean)}")
print()
print(f"Original phone cells     : {original_phone_cells}")
print(f"Kept phone cells         : {kept_phone_cells}")
print(f"Phone cells removed      : {original_phone_cells - kept_phone_cells}")
print()
print(f"Leads with phone         : {leads_with_phone}")
print(f"Leads without phone      : {leads_without_phone}")
print(f"\nCleaned file saved to:\n{out_path}\n")

