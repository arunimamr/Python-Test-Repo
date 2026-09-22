#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\AirCooledChiller"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\AirCooledChiller\\AIR_COOLED_CHILLER.pdf"

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
    "12_EQUIPMENT_OPERATING": "31-181",
    "13_MAINTENANCE_SCOPE": "182-215",
    "14_EQUIPMENT_SETTING": "216-217",
    "15_EQUIPMENT_PASSWORD": "218-219",
    "16_SPARE_PARTS_LIST": "220-221",
    "17_TC__FAT_REPORTS": "224-225",
    "18_TC__L2A_REPORTS": "226-952",
    "19_TC__L2B_REPORTS": "953-1257",
    "20_TC_L3_REPORTS": "1258-1259",
    "21_TC_L4_REPORTS": "1260-1261",
    "22_TC_L5_REPORTS": "1262-1263",
    "23_WARRANTY": "1264-1265",
    "24_MANUFACTURER_PUBLICATION": "1266-2199",
    "25_AS_BUILT_DRAWINGS": "2200-2201",
    "26_CONTRACT_LICENSE": "2202-2204",
    "27_FIRE_CAUSE_EFFECT": "2205-2206",
    "28_LEED": "2207-2208",
    "29_CFD": "2209-2210",
    "30_SPOF": "2211-2212",
    "31_DCOS_POINT_LIST": "2213-2214",
    "32_TRAINING": "2215-2216",
    "33_DEFECT_LIST": "2207-2218",
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
