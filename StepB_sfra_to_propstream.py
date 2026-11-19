import os
from pathlib import Path

import pandas as pd

# ======================================
# Folder paths
# ======================================
base_dir = Path(__file__).resolve().parent
input_folder = base_dir / "StepB - Input SFRA Leads"
output_folder = base_dir / "StepB - Output SFRA Leads in Prop Format"

# Create the pipeline folders before looking for the Step B input CSV.
for folder in (
    input_folder,
    output_folder,
    base_dir / "StepC - Input Prop Skiptrace Export",
    base_dir / "StepC - Input SFRA Leads",
    base_dir / "StepC - Output SFRA Leads w prop Skiptrace",
    base_dir / "StepD - Input SFRA Leads",
    base_dir / "StepD - Input TCPA Liti DNC Match",
    base_dir / "StepD - Output SFRA Leads Scrubbed",
):
    folder.mkdir(exist_ok=True)

print("\n=== STEP B, C AND D FOLDERS READY ===")
print(f"Input folder  : {input_folder}")
print(f"Output folder : {output_folder}")

# ======================================
# Locate SFRA input CSV
# ======================================
input_files = [
    f for f in os.listdir(input_folder)
    if f.lower().endswith(".csv")
]

print("\n=== FILES FOUND IN STEP B INPUT FOLDER ===")
for f in os.listdir(input_folder):
    print(" -", f)

print(f"CSV files found: {len(input_files)}")

if len(input_files) == 0:
    raise FileNotFoundError(
        "\n\n"
        "=== NO INPUT CSV FOUND ===\n\n"
        "No CSV files were found inside the Step B input folder:\n\n"
        f"{input_folder}\n\n"
        "Please place exactly ONE SFR Analytics lead CSV file into this "
        "folder.\n\n"
        "Once the CSV file has been copied into the input folder, run the "
        "script again."
    )

if len(input_files) > 1:
    csv_list = "\n".join(
        f" - {filename}" for filename in input_files
    )

    raise ValueError(
        "\n\n"
        "=== MULTIPLE INPUT CSV FILES DETECTED ===\n\n"
        "This script expects exactly ONE SFR Analytics lead CSV file inside "
        "the Step B input folder.\n\n"
        "The following CSV files were detected:\n\n"
        f"{csv_list}\n\n"
        "The script is not sure which CSV file should be used as the input "
        "file, so processing has been stopped.\n\n"
        "Please move or remove the extra CSV files so that only ONE CSV file "
        "remains inside:\n\n"
        f"{input_folder}\n\n"
        "After only one CSV file remains, run the script again."
    )

input_path = os.path.join(input_folder, input_files[0])

print(f"\n=== USING SFRANALYTICS INPUT FILE ===\n{input_path}\n")

# ======================================
# Load SFRA input CSV
# ======================================
df_sfra = pd.read_csv(input_path, dtype=str).fillna("")

# ======================================
# Verify required SFRA headers exist
# ======================================
required_sfra_headers = [
    "Mailing Address",
    "Mailing City",
    "Mailing State",
    "Mailing Zip"
]

missing_headers = [
    header
    for header in required_sfra_headers
    if header not in df_sfra.columns
]

if missing_headers:
    missing_header_text = "\n".join(
        f" - {header}" for header in missing_headers
    )

    raise ValueError(
        "\n\n"
        "=== REQUIRED SFRA HEADERS NOT FOUND ===\n\n"
        "The input CSV is missing one or more columns required for the "
        "PropStream conversion:\n\n"
        f"{missing_header_text}\n\n"
        "Please confirm that the CSV came from SFR Analytics and contains "
        "the following exact headers:\n\n"
        " - Mailing Address\n"
        " - Mailing City\n"
        " - Mailing State\n"
        " - Mailing Zip\n\n"
        f"Input file:\n{input_path}"
    )

# ======================================
# PropStream import headers
# ======================================
propstream_headers = [
    "ID",
    "Code",
    "Phone Number",
    "Title",
    "First Name",
    "Middle Name",
    "Last Name",
    "Address 1",
    "Address 2",
    "Address 3",
    "City",
    "State",
    "Zip",
    "Country",
    "Gender",
    "Birth Date",
    "Alt Phone Number",
    "Email"
]

# ======================================
# Create PropStream output DataFrame
# ======================================
df_output = pd.DataFrame(columns=propstream_headers)

# ======================================
# Map SFRA fields to PropStream fields
# ======================================
df_output["Address 1"] = df_sfra["Mailing Address"]
df_output["City"] = df_sfra["Mailing City"]
df_output["State"] = df_sfra["Mailing State"]
df_output["Zip"] = df_sfra["Mailing Zip"]

# ======================================
# Build output file name
# ======================================
base, ext = os.path.splitext(input_files[0])

output_filename = base + "_prop_format" + ext
output_path = os.path.join(output_folder, output_filename)

# ======================================
# Save output
# ======================================
try:
    df_output.to_csv(output_path, index=False, mode="x")
except FileExistsError:
    raise FileExistsError(
        f"Step B skipped the export because {output_path} already exists. "
        "Move or rename that file to generate a fresh export."
    ) from None

# ======================================
# Final report
# ======================================
print("\n=== PROPSTREAM IMPORT FILE CREATED ===")
print(f"Input file used : {input_files[0]}")
print(f"Headers created : {len(propstream_headers)}")
print(f"Data rows       : {len(df_output)}")

print(f"\nOutput file saved to:\n{output_path}\n")