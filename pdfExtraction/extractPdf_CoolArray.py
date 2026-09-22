#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\CoolArray"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\CoolArray\\COOL_ARRAY.pdf"

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
    "10_EQUIPMENT_LIST": "21-23",
    "11_DESIGN_CRITERIA": "24-31",
    "12_EQUIPMENT_OPERATING": "32-130",
    "13_MAINTENANCE_SCOPE": "131-151",
    "14_EQUIPMENT_SETTING": "152-153",
    "15_EQUIPMENT_PASSWORD": "154-155",
    "16_SPARE_PARTS_LIST": "156-157",
    "17_TC__FAT_REPORTS": "160-161",
    "18_TC__L2A_REPORTS": "162-7049",
    "19_TC__L2B_REPORTS": "7050-7051",
    "20_TC_L3_REPORTS": "7052-7053",
    "21_TC_L4_REPORTS": "7054-7055",
    "22_TC_L5_REPORTS": "7056-7057",
    "23_WARRANTY": "7058-7059",
    "24_MANUFACTURER_PUBLICATION": "7060-7219",
    "25_AS_BUILT_DRAWINGS": "7220-7221",
    "26_CONTRACT_LICENSE": "7222-7227",
    "27_FIRE_CAUSE_EFFECT": "7228-7229",
    "28_LEED": "7230-7231",
    "29_CFD": "7232-7233",
    "30_SPOF": "7234-7235",
    "31_DCOS_POINT_LIST": "7236-7237",
    "32_TRAINING": "7238-7239",
    "33_DEFECT_LIST": "7240-7241",
}

# === Setup output directory ===
today_folder = datetime.now().strftime("Extracted_%Y_%m_%d")
today = os.path.join(base_dir, today_folder)
os.makedirs(today, exist_ok=True)
print(f'Created directory: {today}')

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
    out_file = os.path.join(today, f"{name}.pdf")
    new_pdf.save(out_file)
    new_pdf.close()
    print(f"  Saved: {out_file}")

doc.close()
print("✅ Extraction completed successfully.")
