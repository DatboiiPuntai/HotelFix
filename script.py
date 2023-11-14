import csv
import sys
import os

import pandas as pd


SELECTED_COLUMNS = [
    "First Name\nชื่อ",
    "Middle Name\nชื่อกลาง",
    "Last Name\nนามสกุล",
    "Sex\nเพศ\n(M-ชาย,F-หญิง)",
    "Passport No.\nหนังสือเดินทาง",
    "Nationality\nสัญชาติ",
    "Date of Birth (mm/dd/yyyy) (A.D.)\nวันที่เกิด (ค.ศ.)",
    "Expire Date of Stay (mm/dd/yyyy) (A.D.)\nวันครบกำหนดอนุญาต (ค.ศ.)",
]

RENAME_MAPPING = {
    "First Name\nชื่อ": "ชื่อ\nFirst Name *",
    "Middle Name\nชื่อกลาง": "ชื่อกลาง\nMiddle Name",
    "Last Name\nนามสกุล": "นามสกุล\nLast Name",
    "Sex\nเพศ\n(M-ชาย,F-หญิง)": "เพศ\nGender *",
    "Passport No.\nหนังสือเดินทาง": "เลขหนังสือเดินทาง\nPassport No. *",
    "Nationality\nสัญชาติ": "สัญชาติ\nNationality *",
    "Date of Birth (mm/dd/yyyy) (A.D.)\nวันที่เกิด (ค.ศ.)": "วัน เดือน ปี เกิด\nBirth Date *\nDD/MM/YYYY(ค.ศ. / A.D.)",
    "Expire Date of Stay (mm/dd/yyyy) (A.D.)\nวันครบกำหนดอนุญาต (ค.ศ.)": "วันที่แจ้งออกจากที่พัก\nCheck-out Date *\nDD/MM/YYYY(ค.ศ. / A.D.)",
}


def read_nationality_codes(csv_file_path):
    """
    Reads nationality codes from a CSV file and returns a dictionary.

    Parameters:
    - csv_file_path (str): Path to the CSV file.

    Returns:
    - dict: Dictionary mapping nationality codes to ICAO codes.
    """
    nationality_dict = {}
    with open(csv_file_path, newline="", encoding="utf-8") as csvfile:
        csv_reader = csv.DictReader(csvfile)
        for row in csv_reader:
            nationality_dict[row["รหัสสัญชาติ"]] = row["รหัส icao"]
    return nationality_dict


def process_and_save_excel(input_filename, output_filename):
    # Read data from the input file
    df = pd.read_excel(input_filename)

    # Selecting and renaming columns
    df = df[SELECTED_COLUMNS]

    # Renaming columns
    df = df.rename(columns=RENAME_MAPPING)

    # Add a new column 'เบอร์โทรศัพท์\nPhone No.' with NaN values
    df["เบอร์โทรศัพท์\nPhone No."] = ""

    df["สัญชาติ\nNationality *"] = df["สัญชาติ\nNationality *"].apply(
        lambda x: nationality_dict.get(x, None)
    )

    # Split the input filename into root and extension
    root, ext = os.path.splitext(input_filename)

    # Determine the output filename
    if output_filename is None:
        output_filename = root + "_processed.xlsx"

    # Save the processed data to a new Excel file
    df.to_excel(output_filename, index=False, sheet_name="แบบแจ้งที่พัก Inform Accom")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_filename> [output_filename]")
        sys.exit(1)

    input_filename = sys.argv[1]

    # Output filename is an optional argument
    output_filename = sys.argv[2] if len(sys.argv) > 2 else None

    nationality_dict = read_nationality_codes("data/config/nationality_code.csv")
    process_and_save_excel(input_filename, output_filename)
