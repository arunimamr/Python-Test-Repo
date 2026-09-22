#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\LIB"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\LIB\\LIB.pdf"

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
    "10_EQUIPMENT_LIST": "21-24",
    "11_DESIGN_CRITERIA": "25-35",
    "12_EQUIPMENT_OPERATING": "36-61",
    "13_MAINTENANCE_SCOPE": "62-72",
    "14_EQUIPMENT_SETTING": "73-74",
    "15_EQUIPMENT_PASSWORD": "75-76",
    "16_SPARE_PARTS_LIST": "77-78",
    "17_TC__FAT_REPORTS": "81-82",
    "18_TC__L2A_REPORTS": "83-9293",
    "19_TC__L2B_REPORTS": "9294-11928",
    "20_TC_L3_REPORTS": "11929-12963",
    "21_TC_L4_REPORTS": "12964-12965",
    "22_TC_L5_REPORTS": "12966-12967",
    "23_WARRANTY": "12968-12969",
    "24_MANUFACTURER_PUBLICATION": "12970-13496",
    "25_AS_BUILT_DRAWINGS": "13497-13498",
    "26_CONTRACT_LICENSE": "13499-13505",
    "27_FIRE_CAUSE_EFFECT": "13506-13507",
    "28_LEED": "13508-13509",
    "29_CFD": "13510-13511",
    "30_SPOF": "13512-13513",
    "31_DCOS_POINT_LIST": "13514-13516",
    "32_TRAINING": "13517-13518",
    "33_DEFECT_LIST": "13519-13520"
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
