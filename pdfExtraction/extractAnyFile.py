import os
import re
from pdf2image import convert_from_path
import pytesseract
from pypdf import PdfReader, PdfWriter

# === CONFIGURATION ===
pdf_path = "E:\\Docs\\ExtractedManuals_Mumbai\\Busduct\\Busduct.pdf"
index_pages = [2, 3]  # pages containing the index
first_content_page = 1  # the real PDF page where page "1" in the document starts
output_dir = r"E:\Docs\ExtractedManuals_Mumbai\Busduct\sections/"

os.makedirs(output_dir, exist_ok=True)

# === STEP 1: Extract text from index pages (OCR if needed) ===
index_text = ""
reader = PdfReader(pdf_path)

for i in index_pages:
    page = reader.pages[i - 1]
    text = page.extract_text()
    if text and len(text.strip()) > 10:
        index_text += "\n" + text
    else:
        images = convert_from_path(pdf_path, first_page=i, last_page=i)
        for img in images:
            text_ocr = pytesseract.image_to_string(img, lang="eng")
            index_text += "\n" + text_ocr

# === STEP 2: Parse titles and logical (document) page numbers ===
# Example line: "Chapter 1 ............. 1"
pattern = re.compile(r"(.+?)\s+(\d+)\s*$", re.MULTILINE)
matches = pattern.findall(index_text)

sections = []
for title, doc_page in matches:
    title = re.sub(r'[\\/*?:"<>|]', "_", title.strip())  # sanitize filename
    sections.append((title, int(doc_page)))

# === STEP 3: Map logical page numbers to real PDF pages ===
# Example: If PDF starts at page 5, then doc_page 1 = real_page 5
sections_real = []
for i in range(len(sections)):
    title, doc_start = sections[i]
    real_start = first_content_page + doc_start - 1
    next_real_start = (
        first_content_page + sections[i + 1][1] - 2
        if i + 1 < len(sections)
        else len(reader.pages)
    )
    sections_real.append((title, real_start, next_real_start))

# === STEP 4: Extract each section ===
for title, start, end in sections_real:
    if start > len(reader.pages):  # skip invalid mapping
        continue
    end = min(end, len(reader.pages))
    writer = PdfWriter()
    for p in range(start - 1, end):
        writer.add_page(reader.pages[p])
    out_path = os.path.join(output_dir, f"{title}.pdf")
    with open(out_path, "wb") as f:
        writer.write(f)
    print(f"✅ Saved: {out_path} ({start}-{end})")

print("🎉 Done! All sections extracted using index names.")
