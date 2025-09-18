## created by Er. Ajay Bhattarai ##
### This is useful for the margin control in the pdf, but this reduces the text horizontal spacing, and text width ###

import fitz  # PyMuPDF
import os
from PIL import Image
import io

def adjust_pdf_margins_image_compression(input_path, output_path, target_margin_cm=2.0, pages_to_process=None):
    """
    Adjusts PDF margins by converting pages to images, applying horizontal compression,
    and recreating the PDF. This approach treats the entire page like Photoshop.
    
    Args:
        input_path (str): Path to input PDF
        output_path (str): Path to output PDF
        target_margin_cm (float): Target margin in centimeters (default: 2.0)
        pages_to_process (list): List of page numbers to process (1-indexed). If None, no pages are processed.
    """
    
    # Convert cm to points (1 cm = 28.35 points)
    target_margin_pt = target_margin_cm * 28.35
    
    # If no pages specified, don't process anything
    if not pages_to_process:
        print("No pages specified for processing. Copying original PDF...")
        # Simply copy the original file
        import shutil
        shutil.copy2(input_path, output_path)
        print(f"Original PDF copied to: {output_path}")
        return
    
    # Convert to set for faster lookup (convert to 0-indexed)
    process_pages_set = set(page - 1 for page in pages_to_process if isinstance(page, int) and page > 0)
    
    try:
        # Open the PDF
        input_doc = fitz.open(input_path)
        print(f"Processing PDF: {input_path}")
        print(f"Total pages: {input_doc.page_count}")
        print(f"Pages to process: {sorted(pages_to_process)}")
        
        # Create new output document
        output_doc = fitz.open()
        
        # Process each page
        for page_num in range(input_doc.page_count):
            page = input_doc[page_num]
            
            # Check if this page should be processed
            if page_num not in process_pages_set:
                print(f"Page {page_num + 1}: SKIPPED (not in processing list)")
                # Copy original page without modification
                output_doc.insert_pdf(input_doc, from_page=page_num, to_page=page_num)
                continue
            
            # Get page dimensions
            page_rect = page.rect
            page_width = page_rect.width
            page_height = page_rect.height
            
            print(f"Page {page_num + 1}: {page_width:.1f} x {page_height:.1f} pts - PROCESSING...")
            
            try:
                # Convert page to high-resolution image (like taking a screenshot)
                dpi = 300  # High resolution for quality
                mat = fitz.Matrix(dpi/72, dpi/72)  # Scale matrix for DPI
                pix = page.get_pixmap(matrix=mat, alpha=False)
                
                # Convert to PIL Image
                img_data = pix.tobytes("png")
                original_image = Image.open(io.BytesIO(img_data))
                
                print(f"  Original image size: {original_image.size}")
                
                # Detect content boundaries in the image
                content_bounds = detect_content_boundaries(original_image, target_margin_cm, page_width)
                
                if not content_bounds:
                    print(f"  No content detected, copying original page...")
                    output_doc.insert_pdf(input_doc, from_page=page_num, to_page=page_num)
                    continue
                
                left_margin_px, right_margin_px, content_width_px = content_bounds
                
                # Calculate current margins in cm
                pixels_per_cm = dpi / 2.54
                current_left_margin_cm = left_margin_px / pixels_per_cm
                current_right_margin_cm = right_margin_px / pixels_per_cm
                
                print(f"  Current margins: Left={current_left_margin_cm:.1f}cm, Right={current_right_margin_cm:.1f}cm")
                
                # Check if adjustment is needed
                if abs(current_left_margin_cm - target_margin_cm) < 0.2:  # 2mm tolerance
                    print(f"  Margins already close to {target_margin_cm}cm, copying original...")
                    output_doc.insert_pdf(input_doc, from_page=page_num, to_page=page_num)
                    continue
                
                # Apply horizontal compression like Photoshop
                compressed_image = apply_horizontal_compression_to_image(
                    original_image, target_margin_cm, dpi
                )
                
                if compressed_image:
                    # Convert compressed image back to PDF page
                    success = insert_image_as_pdf_page(output_doc, compressed_image, page_width, page_height)
                    
                    if success:
                        print(f"  ✓ Page {page_num + 1} successfully compressed with {target_margin_cm}cm margins")
                    else:
                        print(f"  ✗ Failed to insert compressed image, copying original...")
                        output_doc.delete_page(-1)  # Remove failed page
                        output_doc.insert_pdf(input_doc, from_page=page_num, to_page=page_num)
                else:
                    print(f"  ✗ Compression failed, copying original...")
                    output_doc.insert_pdf(input_doc, from_page=page_num, to_page=page_num)
                
                # Clean up
                pix = None
                original_image.close()
                if compressed_image:
                    compressed_image.close()
                
            except Exception as e:
                print(f"  ✗ Error processing page {page_num + 1}: {e}")
                print(f"    Copying original page as fallback...")
                output_doc.insert_pdf(input_doc, from_page=page_num, to_page=page_num)
        
        # Save the output document
        output_doc.save(output_path)
        output_doc.close()
        input_doc.close()
        
        print(f"\n✓ PDF successfully saved as: {output_path}")
        print(f"Processed pages now have {target_margin_cm}cm margins on both sides.")
        
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        if 'input_doc' in locals():
            input_doc.close()
        if 'output_doc' in locals():
            output_doc.close()

def detect_content_boundaries(image, target_margin_cm, page_width_pt):
    """
    Detect content boundaries in the image by finding non-white pixels.
    
    Args:
        image: PIL Image object
        target_margin_cm: Target margin in cm
        page_width_pt: Page width in points
        
    Returns:
        tuple: (left_margin_px, right_margin_px, content_width_px) or None
    """
    try:
        # Convert to grayscale for easier analysis
        gray_image = image.convert('L')
        width, height = gray_image.size
        
        # Find left boundary (first non-white column)
        left_boundary = 0
        for x in range(width):
            column_pixels = [gray_image.getpixel((x, y)) for y in range(height)]
            # If any pixel in column is not white-ish (threshold 240)
            if any(pixel < 240 for pixel in column_pixels):
                left_boundary = x
                break
        
        # Find right boundary (last non-white column)
        right_boundary = width - 1
        for x in range(width - 1, -1, -1):
            column_pixels = [gray_image.getpixel((x, y)) for y in range(height)]
            # If any pixel in column is not white-ish (threshold 240)
            if any(pixel < 240 for pixel in column_pixels):
                right_boundary = x
                break
        
        # Calculate margins and content width
        left_margin_px = left_boundary
        right_margin_px = width - right_boundary - 1
        content_width_px = right_boundary - left_boundary + 1
        
        # Validation
        if content_width_px <= 0 or left_boundary >= right_boundary:
            return None
            
        return (left_margin_px, right_margin_px, content_width_px)
        
    except Exception as e:
        print(f"    Content detection error: {e}")
        return None

def apply_horizontal_compression_to_image(image, target_margin_cm, dpi):
    """
    Apply horizontal compression to the entire image like Photoshop's horizontal scaling.
    
    Args:
        image: PIL Image object
        target_margin_cm: Target margin in cm
        dpi: Image DPI
        
    Returns:
        PIL Image: Compressed image or None if failed
    """
    try:
        width, height = image.size
        pixels_per_cm = dpi / 2.54
        target_margin_px = int(target_margin_cm * pixels_per_cm)
        
        # Calculate new content width
        new_content_width = width - (2 * target_margin_px)
        
        if new_content_width <= 0:
            print("    Image too narrow for target margins")
            return None
        
        # Detect current content boundaries
        content_bounds = detect_content_boundaries(image, target_margin_cm, 0)
        if not content_bounds:
            print("    Could not detect content for compression")
            return None
        
        left_margin_px, right_margin_px, current_content_width = content_bounds
        
        # Calculate compression ratio
        compression_ratio = new_content_width / current_content_width
        
        print(f"    Applying horizontal compression ratio: {compression_ratio:.3f}")
        
        # Extract content area
        content_left = left_margin_px
        content_right = width - right_margin_px
        content_area = image.crop((content_left, 0, content_right, height))
        
        # Compress content horizontally (resize width only)
        compressed_content = content_area.resize((new_content_width, height), Image.Resampling.LANCZOS)
        
        # Create new image with target margins
        new_image = Image.new('RGB', (width, height), 'white')
        
        # Paste compressed content with target left margin
        new_image.paste(compressed_content, (target_margin_px, 0))
        
        return new_image
        
    except Exception as e:
        print(f"    Compression error: {e}")
        return None

def insert_image_as_pdf_page(doc, image, page_width, page_height):
    """
    Insert PIL image as a new PDF page.
    
    Args:
        doc: PyMuPDF document
        image: PIL Image
        page_width: Target page width in points
        page_height: Target page height in points
        
    Returns:
        bool: Success status
    """
    try:
        # Create new page
        new_page = doc.new_page(width=page_width, height=page_height)
        
        # Convert PIL image to bytes
        img_buffer = io.BytesIO()
        image.save(img_buffer, format='PNG', optimize=True, quality=95)
        img_data = img_buffer.getvalue()
        
        # Insert image to fill the entire page
        img_rect = fitz.Rect(0, 0, page_width, page_height)
        new_page.insert_image(img_rect, stream=img_data)
        
        return True
        
    except Exception as e:
        print(f"    PDF insertion error: {e}")
        return False

def get_pages_to_process():
    """
    Get list of pages to process from user input.
    
    Returns:
        list: List of page numbers to process (1-indexed), or None if no pages specified
    """
    print("\nPage Selection for Processing:")
    print("=" * 50)
    print("IMPORTANT: Only specified pages will be processed.")
    print("All other pages will remain unchanged.")
    print("\nEnter page numbers to compress:")
    print("- Single pages: 5,10,25")
    print("- Page ranges: 15-20,35-40")
    print("- Mixed: 5,10,15-20,25,30-35")
    print("- Press Enter without input to process NO pages")
    print("=" * 50)
    
    user_input = input("Pages to compress: ").strip()
    
    if not user_input:
        print("No pages specified. No compression will be applied.")
        return None
    
    pages_to_process = []
    
    try:
        parts = user_input.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                # Handle ranges like "15-20"
                start, end = map(int, part.split('-'))
                pages_to_process.extend(range(start, end + 1))
            else:
                # Handle single page numbers
                pages_to_process.append(int(part))
        
        # Remove duplicates and sort
        pages_to_process = sorted(list(set(pages_to_process)))
        
        return pages_to_process
        
    except ValueError:
        print("Error: Invalid input format. No pages will be processed.")
        return None

def main():
    # File paths
    input_file = r"C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\TUTORIAL\FINAL-PRINT.pdf"
    output_file = r"C:\Users\ajayb\OneDrive - Tribhuvan University\Desktop\TUTORIAL\FINAL-CORRECT.pdf"
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file not found: {input_file}")
        return
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)
    
    print("PDF HORIZONTAL COMPRESSION TOOL")
    print("Photoshop-Style Whole Page Compression")
    print("=" * 70)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print(f"Target margin: 2.0 cm on both sides")
    print("=" * 70)
    print("\nHow this works:")
    print("✓ Converts each specified page to high-resolution image")
    print("✓ Detects content boundaries automatically") 
    print("✓ Applies horizontal compression to entire page content")
    print("✓ Creates exactly 2cm margins on left and right")
    print("✓ Maintains original quality and aspect ratios")
    print("✓ Leaves unspecified pages completely unchanged")
    
    # Get pages to process from user
    pages_to_process = get_pages_to_process()
    
    if pages_to_process:
        print(f"\nPages selected for compression: {pages_to_process}")
        print(f"All other pages will remain unchanged.")
        print("=" * 70)
        
        # Confirm before processing
        confirm = input("Proceed with compression? (y/n): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("Operation cancelled.")
            return
    else:
        print("\nNo pages selected. Creating copy of original PDF...")
    
    print("\nStarting processing...")
    
    # Process the PDF
    adjust_pdf_margins_image_compression(input_file, output_file, target_margin_cm=2.0, pages_to_process=pages_to_process)

if __name__ == "__main__":
    main()