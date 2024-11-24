import fitz  # PyMuPDF
from docx2pdf import convert
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from io import BytesIO
from PIL import Image as PILImage
import tempfile
import os
import re
from docx import Document  # New import to work with DOCX files

def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convert DOCX to PDF using docx2pdf.
    """
    try:
        convert(docx_path, pdf_path)
        print(f"PDF successfully created at {pdf_path}")
    except Exception as e:
        print(f"Error in converting DOCX to PDF: {e}")

def dpi_adjust_image(image_path, target_dpi=300):
    """
    Adjust the image to 300 DPI and return the path to the adjusted image.
    """
    try:
        img = PILImage.open(image_path)
        img = img.convert("RGB")
        
        # Create a temporary file to save the adjusted image
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_path = temp_file.name
        
        # Save the image with the target DPI
        img.save(temp_path, dpi=(target_dpi, target_dpi))
        
        print(f"Image adjusted and saved at {temp_path}")
        return temp_path  # Return the path to the adjusted image
    except Exception as e:
        print(f"Error adjusting DPI for image {image_path}: {e}")
        return None

def find_image_placeholder_position_in_docx(docx_path, placeholder="<Image1>"):
    """
    Find the position of the <Image1> placeholder word-by-word in the DOCX file.
    This function returns the index of the paragraph containing the placeholder.
    """
    doc = Document(docx_path)
    for para_index, paragraph in enumerate(doc.paragraphs):
        words = paragraph.text.split()  # Split the paragraph into words
        for word_index, word in enumerate(words):
            if word == placeholder:
                print(f"Found placeholder '{placeholder}' at word {word_index} in paragraph {para_index}.")
                return para_index  # Return paragraph index where placeholder is found
    return None


def add_images_and_footnote(input_pdf, output_pdf, footnote_text, header_image_path=None, image_path=None, docx_path=None):
    """
    Directly place Image1 in the placeholder <Image1> and add a header image and footnote to the PDF.
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    width, height = A4  # Dimensions of A4 paper

    # Define light blue color using RGB
    light_blue = colors.Color(0.678, 0.847, 0.902)  # RGB for light bluish color

    # Get the placeholder position from DOCX if provided
    placeholder_paragraph_index = find_image_placeholder_position_in_docx(docx_path)

    for page_num, page in enumerate(reader.pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        # Draw the header color band (top band)
        header_height = 50  # Height of the header color band
        can.setFillColor(light_blue)
        can.rect(0, height - header_height, width, header_height, fill=True, stroke=False)  # Full width of A4

        # Add header image if provided
        if header_image_path:
            try:
                header_width = 200  # Adjust width as needed
                header_height_image = 100  # Adjust height as needed
                can.drawImage(
                    header_image_path,
                    60,  # Left margin for the header image
                    height - header_height_image - 5,  # Position near the top (20 pts margin)
                    width=header_width,
                    height=header_height_image,
                    preserveAspectRatio=True,
                )
                print(f"Header image added to page {page_num}.")
            except Exception as e:
                print(f"Error adding header image: {e}")

        if placeholder_paragraph_index is not None and image_path:
            # Adjust the image to 300 DPI
            adjusted_image_path = dpi_adjust_image(image_path)
            if adjusted_image_path:
                try:
                    # Save the PIL image to a temporary file and then use that file for drawing
                    img = PILImage.open(adjusted_image_path)
                    img.save("temp_image.png")

                    # Define the target width and height in inches (5x2 inches)
                    target_width_in = 5  # inches
                    target_height_in = 2  # inches

                    # Convert inches to points (1 inch = 72 points)
                    target_width_pts = target_width_in * 72  # 360 points
                    target_height_pts = target_height_in * 72  # 144 points

                    # Calculate position based on the paragraph index (simplified approach)
                    y_position = height - (placeholder_paragraph_index + 1) * 100  # Simplified logic

                    # Draw the image at the mapped position
                    can.drawImage("temp_image.png", 100, y_position, width=target_width_pts, height=target_height_pts, preserveAspectRatio=True)
                    print(f"Image added at position (100, {y_position}) on page {page_num}.")
                except Exception as e:
                    print(f"Error adding image: {e}")

        # Draw the footer color band (bottom band)
        footer_height = 40  # Height of the footer color band
        can.setFillColor(light_blue)
        can.rect(0, 0, width, footer_height, fill=True, stroke=False)  # Full width of A4

        # Set the fill color to black for the footnote text
        can.setFillColor(colors.black)

        # Add footnote text inside the footer color band
        can.setFont("Times-Italic", 8)
        x_position = 50  # Left margin for footnote
        y_position = 30  # Position inside the footer color band (adjust to fit inside footer)

        # Adjust the vertical positioning of the footnote text inside the footer
        for line in footnote_text.split("\n"):
            can.drawString(x_position, y_position, line.strip())
            y_position -= 10  # Line spacing for the footnote

        can.save()

        # Merge the overlay with the existing page
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    # Save the final output
    with open(output_pdf, "wb") as f:
        writer.write(f)
    print(f"Final PDF with image and footnotes saved at {output_pdf}")


# File paths
docx_path = "paper4.docx"  # Input Word document
temp_pdf_path = "temp_output.pdf"  # Temporary PDF file
final_pdf_path = "formatted_paper_with_images_and_footnotes.pdf"  # Final PDF with images and footnotes
header_image_path = "picture1.png"  # Header image file path
image_path = "Image1.png"  # Image file to be added at <Image1>
footnote_text = (
    "Global Health Equity. © 2024. Published by Global Health Equity journal. Global Health Equity is an Open Access journal distributed, \n"
    "under the terms of the Creative Commons Attribution License (http://creativecommons.org/licenses/by/4.0/), which permits unrestricted use,\n" 
    "distribution, and reproduction in any medium, provided the original work is properly cited."
)

# Ensure the input DOCX exists
if os.path.exists(docx_path):
    # Step 1: Convert DOCX to PDF
    convert_docx_to_pdf(docx_path, temp_pdf_path)

    # Step 2: Add image and footnote to the PDF
    if os.path.exists(temp_pdf_path):
        add_images_and_footnote(temp_pdf_path, final_pdf_path, footnote_text, header_image_path=header_image_path, image_path=image_path, docx_path=docx_path)
    else:
        print(f"Error: {temp_pdf_path} not found.")
else:
    print(f"Error: {docx_path} not found.")
