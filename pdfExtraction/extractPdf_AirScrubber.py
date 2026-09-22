#pip install pymupdf

import fitz  # PyMuPDF
import os
from datetime import datetime

# === Configuration ===

base_dir = r"E:\Docs\ExtractedManuals_Mumbai\AirScrubber"
FILE = "E:\\Docs\\ExtractedManuals_Mumbai\\AirScrubber\\AirScrubber.pdf"

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
    "11_DESIGN_CRITERIA": "24-28",
    "12_EQUIPMENT_OPERATING": "29-41",
    "13_MAINTENANCE_SCOPE": "42-45",
    "14_EQUIPMENT_SETTING": "46-47",
    "15_EQUIPMENT_PASSWORD": "48-49",
    "16_SPARE_PARTS_LIST": "50-51",
    "17_TC__FAT_REPORTS": "54-55",
    "18_TC__L2A_REPORTS": "56-2701",
    "19_TC__L2B_REPORTS": "2702-3563",
    "20_TC_L3_REPORTS": "3564-4448",
    "21_TC_L4_REPORTS": "4449-4450",
    "22_TC_L5_REPORTS": "4451-4452",
    "23_WARRANTY": "4453-4454",
    "24_MANUFACTURER_PUBLICATION": "4455-4460",
    "25_AS_BUILT_DRAWINGS": "4461-4462",
    "26_CONTRACT_LICENSE": "4463-4469",
    "27_FIRE_CAUSE_EFFECT": "4470-4471",
    "28_LEED": "4472-4473",
    "29_CFD": "4474-4475",
    "30_SPOF": "4476-4477",
    "31_DCOS_POINT_LIST": "4478-4479",
    "32_TRAINING": "4480-4481",
    "33_DEFECT_LIST": "4482-4483",
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
