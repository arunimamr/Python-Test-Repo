#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\DBs"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\DBs\\DB.pdf"

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
    "10_EQUIPMENT_LIST": "21-30",
    "11_DESIGN_CRITERIA": "31-37",
    "12_EQUIPMENT_OPERATING": "38-43",
    "13_MAINTENANCE_SCOPE": "44-48",
    "14_EQUIPMENT_SETTING": "49-50",
    "15_EQUIPMENT_PASSWORD": "51-52",
    "16_SPARE_PARTS_LIST": "53-54",
    "17_TC__FAT_REPORTS": "57-58",
    "18_TC__L2A_REPORTS": "59-2426",
    "19_TC__L2B_REPORTS": "2427-3833",
    "20_TC_L3_REPORTS": "3834-4456",
    "21_TC_L4_REPORTS": "4457-4458",
    "22_TC_L5_REPORTS": "4459-4460",
    "23_WARRANTY": "4461-4462",
    "24_MANUFACTURER_PUBLICATION": "4463-4571",
    "25_AS_BUILT_DRAWINGS": "4572-4573",
    "26_CONTRACT_LICENSE": "4574-4582",
    "27_FIRE_CAUSE_EFFECT": "4583-4584",
    "28_LEED": "4585-4586",
    "29_CFD": "4587-4588",
    "30_SPOF": "4589-4590",
    "31_DCOS_POINT_LIST": "4591-4592",
    "32_TRAINING": "4593-4594",
    "33_DEFECT_LIST": "4595-4596",
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
