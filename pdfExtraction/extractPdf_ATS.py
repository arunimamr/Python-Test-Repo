#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\ATS"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\ATS\\ATS.pdf"

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
    "11_DESIGN_CRITERIA": "23-36",
    "12_EQUIPMENT_OPERATING": "37-83",
    "13_MAINTENANCE_SCOPE": "84-100",
    "14_EQUIPMENT_SETTING": "101-102",
    "15_EQUIPMENT_PASSWORD": "103-104",
    "16_SPARE_PARTS_LIST": "105-106",
    "17_TC__FAT_REPORTS": "109-110",
    "18_TC__L2A_REPORTS": "111-326",
    "19_TC__L2B_REPORTS": "327-1179",
    "20_TC_L3_REPORTS": "1180-1181",
    "21_TC_L4_REPORTS": "1182-1183",
    "22_TC_L5_REPORTS": "1184-1185",
    "23_WARRANTY": "1186-1187",
    "24_MANUFACTURER_PUBLICATION": "1188-1224",
    "25_AS_BUILT_DRAWINGS": "1225-1226",
    "26_CONTRACT_LICENSE": "1227-1228",
    "27_FIRE_CAUSE_EFFECT": "1229-1230",
    "28_LEED": "1231-1232",
    "29_CFD": "1233-1234",
    "30_SPOF": "1235-1236",
    "31_DCOS_POINT_LIST": "1237-1238",
    "32_TRAINING": "1239-1240",
    "33_DEFECT_LIST": "1241-1242"
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
