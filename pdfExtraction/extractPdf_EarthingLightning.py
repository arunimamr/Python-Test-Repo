#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\EarthingLightning"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\EarthingLightning\\EarthingLightning.pdf"

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
    "11_DESIGN_CRITERIA": "23-30",
    "12_EQUIPMENT_OPERATING": "31-39",
    "13_MAINTENANCE_SCOPE": "40-43",
    "14_EQUIPMENT_SETTING": "44-45",
    "15_EQUIPMENT_PASSWORD": "46-47",
    "16_SPARE_PARTS_LIST": "48-49",
    "17_TC__FAT_REPORTS": "52-53",
    "18_TC__L2A_REPORTS": "54-55",
    "19_TC__L2B_REPORTS": "56-57",
    "20_TC_L3_REPORTS": "58-59",
    "21_TC_L4_REPORTS": "60-639",
    "22_TC_L5_REPORTS": "62-63",
    "23_WARRANTY": "64-65",
    "24_MANUFACTURER_PUBLICATION": "66-121",
    "25_AS_BUILT_DRAWINGS": "122-123",
    "26_CONTRACT_LICENSE": "124-131",
    "27_FIRE_CAUSE_EFFECT": "132-133",
    "28_LEED": "134-135",
    "29_CFD": "136-137",
    "30_SPOF": "138-139",
    "31_DCOS_POINT_LIST": "140-141",
    "32_TRAINING": "142-143",
    "33_DEFECT_LIST": "144-145",
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
