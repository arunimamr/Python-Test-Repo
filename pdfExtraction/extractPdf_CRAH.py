#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\CRAH"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\CRAH\\CRAH.pdf"

sections = {
    "01_FRONT_PAGE": "1",
    "02_TABLE_OF_CONTENTS": "2-3",
    "03_STAKE_HOLDERS_LIST": "6-9",
    "04_DOCUMENT_CHANGE": "10-11",
    "05_RELATED_DOCUMENT": "12-13",
    "06_ABBREVIATIONS": "14-16",
    "07_INTRODUCTION__HOW_TO_USE": "18",
    "08_INTRODUCTION__DIRECTORY": "19-20",
    "09_INTRODUCTION__EMERGENCY_INFO": "20",
    "10_EQUIPMENT_LIST": "21-22",
    "11_DESIGN_CRITERIA": "23-29",
    "12_EQUIPMENT_OPERATING": "30-101",
    "13_MAINTENANCE_SCOPE": "102-109",
    "14_EQUIPMENT_SETTING": "110-111",
    "15_EQUIPMENT_PASSWORD": "112-113",
    "16_SPARE_PARTS_LIST": "114-115",
    "17_TC__FAT_REPORTS": "118-119",
    "18_TC__L2A_REPORTS": "120-4528",
    "19_TC__L2B_REPORTS": "4529-6034",
    "20_TC_L3_REPORTS": "6035-6036",
    "21_TC_L4_REPORTS": "6037-6038",
    "22_TC_L5_REPORTS": "6039-6040",
    "23_WARRANTY": "6041-6042",
    "24_MANUFACTURER_PUBLICATION": "6043-6210",
    "25_AS_BUILT_DRAWINGS": "6211-6212",
    "26_CONTRACT_LICENSE": "6213-6218",
    "27_FIRE_CAUSE_EFFECT": "6219-6220",
    "28_LEED": "6221-6222",
    "29_CFD": "6223-6224",
    "30_SPOF": "6225-6226",
    "31_DCOS_POINT_LIST": "6227-6228",
    "32_TRAINING": "6229-6230",
    "33_DEFECT_LIST": "6231-6232",
}

# === Setup output directory ===
current_date_folder = datetime.now().strftime("Extracted_%Y_%m_%d")
current_date = os.path.join(base_dir, current_date_folder)
os.makedirs(current_date, exist_ok=True)
print(f'Created directory: {current_date}')

# === Extraction ===
print(f"Extracting from {FILE} ...")
doc = fitz.open(FILE)

for name, page_range in sections.items():
    new_pdf = fitz.open()
    parts = page_range.split('-')

    if len(parts) == 1:
        start = end = int(parts[0]) - 1
    else:
        start, end = int(parts[0]) - 1, int(parts[1]) - 1

    new_pdf.insert_pdf(doc, from_page=start, to_page=end)
    out_file = os.path.join(current_date, f"{name}.pdf")
    new_pdf.save(out_file)
    new_pdf.close()
    print(f"  Saved: {out_file}")

doc.close()
print("✅ Extraction completed successfully.")
