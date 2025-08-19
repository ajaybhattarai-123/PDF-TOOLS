from PyPDF2 import PdfReader, PdfWriter

# Input and output file paths
input_path = r"C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\raga\PRE-PRINT\SONG-FINAL.pdf"
output_path = r"C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\raga\PRE-PRINT\SONG-FINAL-split.pdf"

# Define the range of pages to remove (0-based index)
remove_start = 15 # first page to remove
remove_end = 30    # last page to remove (inclusive)

# Load the PDF
reader = PdfReader(input_path)
writer = PdfWriter()

total_pages = len(reader.pages)

# Add pages to writer, skipping the range to remove
for page_num in range(total_pages):
    if page_num < remove_start or page_num > remove_end:
        writer.add_page(reader.pages[page_num])

# Write to new PDF
with open(output_path, "wb") as output_file:
    writer.write(output_file)

print(f"PDF created successfully without pages {remove_start+1} to {remove_end+1}: {output_path}")
