#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\STS"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\STS\\STS.pdf"

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
    "11_DESIGN_CRITERIA": "23-33",
    "12_EQUIPMENT_OPERATING": "34-83",
    "13_MAINTENANCE_SCOPE": "84-89",
    "14_EQUIPMENT_SETTING": "90-91",
    "15_EQUIPMENT_PASSWORD": "92-93",
    "16_SPARE_PARTS_LIST": "94-95",
    "17_TC__FAT_REPORTS": "98-99",
    "18_TC__L2A_REPORTS": "100-1320",
    "19_TC__L2B_REPORTS": "1321-2550",
    "20_TC_L3_REPORTS": "2551-3627",
    "21_TC_L4_REPORTS": "3628-3629",
    "22_TC_L5_REPORTS": "3630-3631",
    "23_WARRANTY": "3632-3633",
    "24_MANUFACTURER_PUBLICATION": "3634-4016",
    "25_AS_BUILT_DRAWINGS": "4017-4018",
    "26_CONTRACT_LICENSE": "4019-4020",
    "27_FIRE_CAUSE_EFFECT": "4021-4022",
    "28_LEED": "4023-4024",
    "29_CFD": "4025-4026",
    "30_SPOF": "4027-4028",
    "31_DCOS_POINT_LIST": "4029-4030",
    "32_TRAINING": "4031-4032",
    "33_DEFECT_LIST": "4033-4034",
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
