import os
from PyPDF2 import PdfMerger

# Define the folder path
folder_path = r'C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\pdf'

# 👉 SECTION: List the PDF filenames in the exact order you want to merge
pdf_list = [
    "1.pdf",
    "3.pdf",
    "5.pdf",    # Let's say this one is missing
    "7.pdf",
    "15.pdf",
]

# Filter available files
available_files = [pdf for pdf in pdf_list if os.path.exists(os.path.join(folder_path, pdf))]
missing_files = [pdf for pdf in pdf_list if pdf not in available_files]

# Report missing files
if missing_files:
    print("⚠️ These files are missing and will be skipped:")
    for missing in missing_files:
        print(" -", missing)

# Check if anything remains to merge
if not available_files:
    print("❌ No valid PDF files found to merge.")
    exit()

# Define output file name using first and last available filenames
first_name = os.path.splitext(available_files[0])[0]
last_name = os.path.splitext(available_files[-1])[0]
output_filename = f'{first_name}-{last_name}.pdf'
output_path = os.path.join(folder_path, output_filename)

# Merge available PDFs
merger = PdfMerger()

for pdf in available_files:
    merger.append(os.path.join(folder_path, pdf))

# Save the merged PDF
merger.write(output_path)
merger.close()

print(f"\n✅ Merged PDF saved as: {output_path}")
