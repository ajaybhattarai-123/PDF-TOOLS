import os
from PyPDF2 import PdfMerger

# Define the folder path
folder_path = r'C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\pdf'

# 👉 SECTION: List of PDF filenames to skip (include extension)
skip_list = [
    "2.pdf",    # Example: skip 2.pdf
    "10.pdf",   # Add as many as needed
    # "13.pdf",
]

# Get all PDF filenames in the folder, filter and sort them numerically
pdf_files = sorted(
    [f for f in os.listdir(folder_path) if f.endswith('.pdf') and f not in skip_list],
    key=lambda x: int(os.path.splitext(x)[0])
)

# If there are no PDF files, exit
if not pdf_files:
    print("No PDF files found to merge (or all skipped).")
    exit()

# Define output file name using the first and last valid PDF names
first_name = os.path.splitext(pdf_files[0])[0]
last_name = os.path.splitext(pdf_files[-1])[0]
output_filename = f'{first_name}-{last_name}.pdf'
output_path = os.path.join(folder_path, output_filename)

# Merge PDFs
merger = PdfMerger()

for pdf in pdf_files:
    merger.append(os.path.join(folder_path, pdf))

# Write out the merged PDF
merger.write(output_path)
merger.close()

print(f"Merged PDF saved as: {output_path}")
