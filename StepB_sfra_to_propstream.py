import os
import pandas as pd

# ======================================
# Folder paths
# ======================================
base_dir = os.path.dirname(os.path.abspath(__file__))

input_folder = os.path.join(
    base_dir,
    "StepB - Input SFRA Leads"
)

output_folder = os.path.join(
    base_dir,
    "StepB - Output SFRA Leads in Prop Format"
)

# ======================================
# Verify required folders exist
# ======================================
missing_folders = []

if not os.path.isdir(input_folder):
    missing_folders.append(input_folder)

if not os.path.isdir(output_folder):
    missing_folders.append(output_folder)

if missing_folders:
    missing_folder_text = "\n".join(
        f" - {folder}" for folder in missing_folders
    )

    raise FileNotFoundError(
        "\n\n"
        "=== REQUIRED STEP B FOLDER NOT FOUND ===\n\n"
        "The script cannot continue because one or more required folders "
        "do not exist:\n\n"
        f"{missing_folder_text}\n\n"
        "Please manually create each missing folder inside the following "
        "base directory:\n\n"
        f"{base_dir}\n\n"
        "The folder names must match exactly, including all spaces, "
        "capitalization, and punctuation:\n\n"
        "  StepB - Input SFRA Leads\n"
        "  StepB - Output SFRA Leads in Prop Format\n\n"
        "Do not delete or rename these two folders before running the script.\n"
        "The script expects both folders to already exist and will stop if "
        "either folder is missing.\n\n"
        "After creating the folders, place exactly ONE SFR Analytics lead "
        "CSV file inside:\n\n"
        f"{input_folder}\n\n"
        "Then run the script again."
    )

print("\n=== REQUIRED STEP B FOLDERS FOUND ===")
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
        "Do not delete or rename the Step B input or output folders.\n"
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
        "Do not delete or rename the input or output folders themselves.\n"
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
df_output.to_csv(output_path, index=False)

# ======================================
# Final report
# ======================================
print("\n=== PROPSTREAM IMPORT FILE CREATED ===")
print(f"Input file used : {input_files[0]}")
print(f"Headers created : {len(propstream_headers)}")
print(f"Data rows       : {len(df_output)}")

print(f"\nOutput file saved to:\n{output_path}\n")