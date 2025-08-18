from PyPDF2 import PdfReader, PdfWriter

# Input and output file paths
input_path = r"D:\Z-APPLIED MECHANICS\Russell 14th.pdf"
output_path = r"D:\Z-APPLIED MECHANICS\486-528.pdf"

# Define the range (page 220 to 261 inclusive)
start_page = 486 # 486th page (0-based index)
end_page = 528    # 528st page (0-based index)

# Load the PDF
reader = PdfReader(input_path)
writer = PdfWriter()

# Add selected pages
for page_num in range(start_page, end_page):
    writer.add_page(reader.pages[page_num])

# Write to new PDF
with open(output_path, "wb") as output_file:
    writer.write(output_file)

print(f"PDF created successfully: {output_path}")
