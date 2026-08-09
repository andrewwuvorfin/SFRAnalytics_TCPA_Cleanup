import os
import re
import csv
import pandas as pd

# ======================================
# Folder paths
# ======================================
base_dir = r"C:\Users\awuzw\Desktop\Python Scripts\SFRAnalytics_TCPA_Cleanup"

prop_skiptrace_folder = os.path.join(
    base_dir,
    "StepC - Input Prop Skiptrace Export"
)

sfra_input_folder = os.path.join(
    base_dir,
    "StepC - Input SFRA Leads"
)

output_folder = os.path.join(
    base_dir,
    "StepC - Output SFRA Leads w prop Skiptrace"
)

# ======================================
# Normalize address for matching
# ======================================
def normalize_address(value):

    if not value:
        return ""

    address = str(value).upper().strip()

    # Remove punctuation
    address = re.sub(r"[.,]", " ", address)

    # Normalize apartment / unit identifiers
    address = re.sub(r"#", " UNIT ", address)
    address = re.sub(r"\bAPT\b", " UNIT ", address)
    address = re.sub(r"\bAPARTMENT\b", " UNIT ", address)
    address = re.sub(r"\bUNIT\b", " UNIT ", address)

    # Normalize suite identifiers
    address = re.sub(r"\bSUITE\b", " STE ", address)
    address = re.sub(r"\bSTE\b", " STE ", address)

    # Normalize common street suffixes
    replacements = {
        "STREET": "ST",
        "ROAD": "RD",
        "DRIVE": "DR",
        "AVENUE": "AVE",
        "BOULEVARD": "BLVD",
        "COURT": "CT",
        "LANE": "LN",
        "PLACE": "PL",
        "PARKWAY": "PKWY",
        "HIGHWAY": "HWY",
        "CIRCLE": "CIR",
        "TERRACE": "TER",
        "TRAIL": "TRL",
    }

    for old, new in replacements.items():
        address = re.sub(
            rf"\b{old}\b",
            new,
            address
        )

    # Collapse multiple spaces
    address = re.sub(
        r"\s+",
        " ",
        address
    ).strip()

    return address


# ======================================
# Return phone as 10 digits
# ======================================
def phone_digits(value):

    if not value:
        return ""

    digits = re.sub(
        r"\D",
        "",
        str(value)
    )

    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    return digits if len(digits) == 10 else ""


# ======================================
# Format phone number for SFRA output
# ======================================
def normalize_phone(value):

    digits = phone_digits(value)

    if not digits:
        return ""

    return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"


# ======================================
# Verify required folders exist
# ======================================
required_folders = [
    prop_skiptrace_folder,
    sfra_input_folder,
    output_folder
]

missing_folders = [
    folder
    for folder in required_folders
    if not os.path.isdir(folder)
]

if missing_folders:
    raise FileNotFoundError(
        "\n\n=== REQUIRED STEP C FOLDER NOT FOUND ===\n\n"
        "The following required folder(s) do not exist:\n\n"
        + "\n".join(
            f" - {folder}"
            for folder in missing_folders
        )
        + "\n\nPlease manually create the missing folder(s) using the exact "
        "folder names shown above. Do not delete or rename the Step C folders."
    )

# ======================================
# Find input files
# ======================================
prop_files = [
    f
    for f in os.listdir(prop_skiptrace_folder)
    if f.lower().endswith((".xlsx", ".xls", ".csv"))
]

sfra_files = [
    f
    for f in os.listdir(sfra_input_folder)
    if f.lower().endswith(".csv")
]

# ======================================
# Verify exactly one input file per folder
# ======================================
if len(prop_files) != 1:
    raise ValueError(
        "\n\n=== PROPSTREAM INPUT FILE ERROR ===\n\n"
        f"Expected exactly ONE Excel file inside:\n"
        f"{prop_skiptrace_folder}\n\n"
        f"Excel files detected: {len(prop_files)}\n\n"
        "The script is not sure which file should be used. "
        "Leave exactly one PropStream skip-trace export Excel file "
        "in this folder."
    )

if len(sfra_files) != 1:
    raise ValueError(
        "\n\n=== SFRA INPUT FILE ERROR ===\n\n"
        f"Expected exactly ONE CSV file inside:\n"
        f"{sfra_input_folder}\n\n"
        f"CSV files detected: {len(sfra_files)}\n\n"
        "The script is not sure which file should be used. "
        "Leave exactly one SFRA lead CSV file in this folder."
    )

prop_skiptrace_path = os.path.join(
    prop_skiptrace_folder,
    prop_files[0]
)

sfra_input_path = os.path.join(
    sfra_input_folder,
    sfra_files[0]
)

print("\n=== STEP C INPUTS VALIDATED ===")
print(f"PropStream Excel : {prop_skiptrace_path}")
print(f"SFRA CSV         : {sfra_input_path}")
print(f"Output folder    : {output_folder}")

# ======================================
# Load PropStream skip-trace export
# ======================================
df_prop = pd.read_excel(
    prop_skiptrace_path,
    dtype=str
).fillna("")

# ======================================
# Verify required PropStream headers
# ======================================
required_prop_headers = [
    "Address",
    "Phone 1",
    "Phone 2",
    "Phone 3",
    "Phone 4",
    "Phone 5"
]

missing_prop_headers = [
    header
    for header in required_prop_headers
    if header not in df_prop.columns
]

if missing_prop_headers:
    raise ValueError(
        "\n\n=== REQUIRED PROPSTREAM HEADERS NOT FOUND ===\n\n"
        "The following required headers are missing:\n\n"
        + "\n".join(
            f" - {header}"
            for header in missing_prop_headers
        )
        + "\n\nPlease confirm that the Excel file is a PropStream "
        "skip-trace export."
    )

# ======================================
# Build normalized address -> phone list dictionary
# ======================================
prop_phone_dict = {}

prop_phone_columns = [
    "Phone 1",
    "Phone 2",
    "Phone 3",
    "Phone 4",
    "Phone 5"
]

for _, row in df_prop.iterrows():

    raw_address = row.get("Address", "")
    normalized_address = normalize_address(raw_address)

    if not normalized_address:
        continue

    phones = []
    phones_seen = set()

    for phone_column in prop_phone_columns:

        formatted_phone = normalize_phone(
            row.get(phone_column, "")
        )

        digits = phone_digits(formatted_phone)

        if digits and digits not in phones_seen:
            phones.append(formatted_phone)
            phones_seen.add(digits)

    prop_phone_dict[normalized_address] = phones

# ======================================
# Print address -> phone dictionary
# ======================================
print("\n=== PROPSTREAM ADDRESS -> PHONE DICTIONARY ===\n")

for address, phones in prop_phone_dict.items():

    print(address)

    if phones:
        for phone in phones:
            print(f"    {phone}")
    else:
        print("    [NO PHONE NUMBERS]")

    print()

print("=== PROPSTREAM DICTIONARY SUMMARY ===")
print(f"PropStream rows loaded : {len(df_prop)}")
print(f"Addresses added        : {len(prop_phone_dict)}")

# ======================================
# Build output file path
# ======================================
base, ext = os.path.splitext(sfra_files[0])

output_filename = (
    base
    + "_with_prop_skiptrace"
    + ext
)

output_path = os.path.join(
    output_folder,
    output_filename
)

# ======================================
# SFRA phone and new output columns
# ======================================
sfra_phone_columns = [
    "First Phone",
    "Second Phone",
    "Third Phone"
]

new_output_columns = [
    "First Name",
    "Last Name",
    "Propstream Phone 1",
    "Propstream Phone 2",
    "Propstream Phone 3",
    "Propstream Phone 4",
    "Propstream Phone 5"
]

prop_output_phone_columns = [
    "Propstream Phone 1",
    "Propstream Phone 2",
    "Propstream Phone 3",
    "Propstream Phone 4",
    "Propstream Phone 5"
]

# ======================================
# Process SFRA rows and write output
# ======================================
rows_processed = 0
addresses_matched = 0
addresses_not_matched = 0
prop_phones_written = 0
duplicate_prop_phones_skipped = 0

with open(
    sfra_input_path,
    "r",
    newline="",
    encoding="utf-8-sig"
) as input_file:

    reader = csv.DictReader(input_file)

    if reader.fieldnames is None:
        raise ValueError(
            "\n\n=== SFRA HEADER ERROR ===\n\n"
            "The SFRA input CSV does not contain a readable header row."
        )

    required_sfra_headers = [
        "Mailing Address",
        "First Phone",
        "Second Phone",
        "Third Phone"
    ]

    missing_sfra_headers = [
        header
        for header in required_sfra_headers
        if header not in reader.fieldnames
    ]

    if missing_sfra_headers:
        raise ValueError(
            "\n\n=== REQUIRED SFRA HEADERS NOT FOUND ===\n\n"
            "The following required headers are missing:\n\n"
            + "\n".join(
                f" - {header}"
                for header in missing_sfra_headers
            )
        )

    output_headers = reader.fieldnames.copy()

    for new_column in new_output_columns:
        if new_column not in output_headers:
            output_headers.append(new_column)

    with open(
        output_path,
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as output_file:

        writer = csv.DictWriter(
            output_file,
            fieldnames=output_headers
        )

        writer.writeheader()

        for row in reader:

            rows_processed += 1

            # Leave formatting columns blank
            row["First Name"] = ""
            row["Last Name"] = ""

            # Clear the five PropStream output slots
            for column in prop_output_phone_columns:
                row[column] = ""

            original_address = row.get(
                "Mailing Address",
                ""
            )

            normalized_address = normalize_address(
                original_address
            )

            prop_phones = prop_phone_dict.get(
                normalized_address,
                []
            )

            if prop_phones:
                addresses_matched += 1
            else:
                addresses_not_matched += 1

            # Build set of existing SFRA phone numbers
            sfra_phone_digits = set()

            for phone_column in sfra_phone_columns:

                digits = phone_digits(
                    row.get(phone_column, "")
                )

                if digits:
                    sfra_phone_digits.add(digits)

            # Keep only PropStream phones not already in SFRA
            additional_prop_phones = []
            additional_phone_digits = set()

            for prop_phone in prop_phones:

                digits = phone_digits(prop_phone)

                if not digits:
                    continue

                if digits in sfra_phone_digits:
                    duplicate_prop_phones_skipped += 1
                    continue

                if digits in additional_phone_digits:
                    continue

                additional_prop_phones.append(
                    normalize_phone(prop_phone)
                )

                additional_phone_digits.add(digits)

            # Populate available PropStream phone slots
            for index, prop_phone in enumerate(
                additional_prop_phones[:5]
            ):
                output_column = prop_output_phone_columns[index]
                row[output_column] = prop_phone
                prop_phones_written += 1

            writer.writerow(row)

            # Save written rows to disk continuously
            output_file.flush()

# ======================================
# Final report
# ======================================
print("\n=== STEP C OUTPUT FILE CREATED ===")
print(f"SFRA rows processed              : {rows_processed}")
print(f"Addresses matched                : {addresses_matched}")
print(f"Addresses not matched            : {addresses_not_matched}")
print(f"New PropStream phones written    : {prop_phones_written}")
print(f"Existing SFRA duplicates skipped : {duplicate_prop_phones_skipped}")

print(f"\nOutput file saved to:\n{output_path}\n")