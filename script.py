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
    Reads a CSV file containing nationality codes and corresponding ICAO codes.

    Parameters:
    - csv_file_path (str): The path to the CSV file.

    Returns:
    - dict: A dictionary mapping nationality codes to ICAO codes.
    """
    with open(csv_file_path, newline="", encoding="utf-8") as csvfile:
        csv_reader = csv.DictReader(csvfile)
        return {row["รหัสสัญชาติ"]: row["รหัส icao"] for row in csv_reader}


def process_and_save_excel(input_filename, output_filename):
    """
    Processes an Excel file, renames columns, and saves the result to a new Excel file.

    Parameters:
    - input_filename (str): The path to the input Excel file.
    - output_filename (str, optional): The path to the output Excel file. If not provided, a default name is used.

    Returns:
    - None
    """
    df = pd.read_excel(input_filename)[SELECTED_COLUMNS]
    df = df.rename(columns=RENAME_MAPPING)

    df["เบอร์โทรศัพท์\nPhone No."] = ""
    df["สัญชาติ\nNationality *"] = df["สัญชาติ\nNationality *"].map(nationality_dict)

    root, ext = os.path.splitext(input_filename)
    output_filename = output_filename or f"{root}_processed.xlsx"

    df.to_excel(output_filename, index=False, sheet_name="แบบแจ้งที่พัก Inform Accom")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_filename> [output_filename]")
        sys.exit(1)

    input_filename = sys.argv[1]
    output_filename = sys.argv[2] if len(sys.argv) > 2 else None

    nationality_dict = read_nationality_codes("data/config/nationality_code.csv")
    process_and_save_excel(input_filename, output_filename)
