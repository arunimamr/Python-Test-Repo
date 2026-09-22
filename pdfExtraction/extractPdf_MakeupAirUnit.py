#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\MakeupAirUnit"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\MakeupAirUnit\\MakeupAirUnit.pdf"

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
    "11_DESIGN_CRITERIA": "23-31",
    "12_EQUIPMENT_OPERATING": "32-41",
    "13_MAINTENANCE_SCOPE": "42-49",
    "14_EQUIPMENT_SETTING": "50-51",
    "15_EQUIPMENT_PASSWORD": "52-53",
    "16_SPARE_PARTS_LIST": "54-55",
    "17_TC__FAT_REPORTS": "58-59",
    "18_TC__L2A_REPORTS": "60-137",
    "19_TC__L2B_REPORTS": "138-343",
    "20_TC_L3_REPORTS": "344-345",
    "21_TC_L4_REPORTS": "346-347",
    "22_TC_L5_REPORTS": "348-349",
    "23_WARRANTY": "350-351",
    "24_MANUFACTURER_PUBLICATION": "352-432",
    "25_AS_BUILT_DRAWINGS": "433-434",
    "26_CONTRACT_LICENSE": "435-438",
    "27_FIRE_CAUSE_EFFECT": "439-440",
    "28_LEED": "441-442",
    "29_CFD": "443-444",
    "30_SPOF": "445-446",
    "31_DCOS_POINT_LIST": "447-448",
    "32_TRAINING": "449-450",
    "33_DEFECT_LIST": "451-452"
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
