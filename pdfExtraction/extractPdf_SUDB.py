#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\SUDB"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\SUDB\\SUDB.pdf"

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
    "11_DESIGN_CRITERIA": "23-49",
    "12_EQUIPMENT_OPERATING": "50-108",
    "13_MAINTENANCE_SCOPE": "109-117",
    "14_EQUIPMENT_SETTING": "118-325",
    "15_EQUIPMENT_PASSWORD": "326-327",
    "16_SPARE_PARTS_LIST": "328-329",
    "17_TC__FAT_REPORTS": "332-416",
    "18_TC__L2A_REPORTS": "417-807",
    "19_TC__L2B_REPORTS": "808-1022",
    "20_TC_L3_REPORTS": "1023-1147",
    "21_TC_L4_REPORTS": "1148-1149",
    "22_TC_L5_REPORTS": "1150-1151",
    "23_WARRANTY": "1152-1153",
    "24_MANUFACTURER_PUBLICATION": "1154-1593",
    "25_AS_BUILT_DRAWINGS": "1594-1595",
    "26_CONTRACT_LICENSE": "1596-1600",
    "27_FIRE_CAUSE_EFFECT": "1601-1602",
    "28_LEED": "1603-1604",
    "29_CFD": "1605-1606",
    "30_SPOF": "1607-1608",
    "31_DCOS_POINT_LIST": "1609-1610",
    "32_TRAINING": "1611-1612",
    "33_DEFECT_LIST": "1613-1614",
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
