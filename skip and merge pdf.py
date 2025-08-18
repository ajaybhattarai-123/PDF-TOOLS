import os
from PyPDF2 import PdfReader, PdfWriter

# 👉 Folder where PDFs are stored
folder_path = r'C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\pdf'

# 👉 Dictionary: PDF filename -> list of pages to skip (1-based)
pdf_skip_dict = {
    "1.pdf": [2, 7, 50],
    "2.pdf": [1, 25],
    "3.pdf": [],
    "4.pdf": [1, 2],  # Skipping first 2 pages
}

# Filter existing files and track missing
existing_dict = {k: v for k, v in pdf_skip_dict.items() if os.path.exists(os.path.join(folder_path, k))}
missing_files = [k for k in pdf_skip_dict if k not in existing_dict]

if missing_files:
    print("⚠️ Missing files (skipped):")
    for f in missing_files:
        print(f" - {f}")

if not existing_dict:
    print("❌ No valid PDF files to process.")
    exit()

# Get first and last filenames for output name
valid_keys = list(existing_dict.keys())
first_name = os.path.splitext(valid_keys[0])[0]
last_name = os.path.splitext(valid_keys[-1])[0]
output_filename = f"{first_name}-{last_name}.pdf"
output_path = os.path.join(folder_path, output_filename)

# Prepare writer
writer = PdfWriter()

# Loop through each file and skip specified pages
for pdf_name, skip_pages in existing_dict.items():
    pdf_path = os.path.join(folder_path, pdf_name)
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    
    # Convert 1-based skip list to 0-based index
    skip_set = set([p - 1 for p in skip_pages if 1 <= p <= total_pages])

    for i in range(total_pages):
        if i not in skip_set:
            writer.add_page(reader.pages[i])

# Save final merged PDF
with open(output_path, 'wb') as f:
    writer.write(f)

print(f"\n✅ Merged PDF saved as: {output_path}")
