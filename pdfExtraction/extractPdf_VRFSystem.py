#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\VRFSystem"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\VRFSystem\\VRFSystem.pdf"

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
    "10_EQUIPMENT_LIST": "21-32",
    "11_DESIGN_CRITERIA": "33-66",
    "12_EQUIPMENT_OPERATING": "67-98",
    "13_MAINTENANCE_SCOPE": "99-106",
    "14_EQUIPMENT_SETTING": "107-108",
    "15_EQUIPMENT_PASSWORD": "109-110",
    "16_SPARE_PARTS_LIST": "111-112",
    "17_TC__FAT_REPORTS": "115-116",
    "18_TC__L2A_REPORTS": "117-5667",
    "19_TC__L2B_REPORTS": "5668-14948",
    "20_TC_L3_REPORTS": "14949-14950",
    "21_TC_L4_REPORTS": "14951-14952",
    "22_TC_L5_REPORTS": "14953-14954",
    "23_WARRANTY": "14955-14956",
    "24_MANUFACTURER_PUBLICATION": "14957-15379",
    "25_AS_BUILT_DRAWINGS": "15380-15381",
    "26_CONTRACT_LICENSE": "15382-15390",
    "27_FIRE_CAUSE_EFFECT": "15391-15392",
    "28_LEED": "15393-15394",
    "29_CFD": "15395-15396",
    "30_SPOF": "15397-15398",
    "31_DCOS_POINT_LIST": "15399-15400",
    "32_TRAINING": "15401-15402",
    "33_DEFECT_LIST": "15403-15404",
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
