from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
import tempfile
from docx2pdf import convert
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from io import BytesIO
from PIL import Image as PILImage
from docx import Document

# Create Flask app
app = Flask(__name__)

# Configure file upload
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx', 'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the uploads directory exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Helper function to check allowed file extensions
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Convert DOCX to PDF using docx2pdf
def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        convert(docx_path, pdf_path)
        return pdf_path
    except Exception as e:
        print(f"Error in converting DOCX to PDF: {e}")
        return None

# Adjust image DPI to 300
def dpi_adjust_image(image_path, target_dpi=300):
    try:
        img = PILImage.open(image_path)
        img = img.convert("RGB")
        
        # Create a temporary file to save the adjusted image
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_path = temp_file.name
        
        # Save the image with the target DPI
        img.save(temp_path, dpi=(target_dpi, target_dpi))
        return temp_path
    except Exception as e:
        print(f"Error adjusting DPI for image {image_path}: {e}")
        return None

# Find placeholder position in DOCX file
def find_image_placeholder_position_in_docx(docx_path, placeholder="<Image1>"):
    doc = Document(docx_path)
    for para_index, paragraph in enumerate(doc.paragraphs):
        words = paragraph.text.split()
        for word_index, word in enumerate(words):
            if word == placeholder:
                return para_index  # Return paragraph index
    return None

# Add image and footnotes to the PDF
def add_images_and_footnote(input_pdf, output_pdf, footnote_text, header_image_path=None, image_path=None, docx_path=None):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    width, height = A4

    light_blue = colors.Color(0.678, 0.847, 0.902)

    placeholder_paragraph_index = find_image_placeholder_position_in_docx(docx_path)

    for page_num, page in enumerate(reader.pages):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        # Draw header
        header_height = 50
        can.setFillColor(light_blue)
        can.rect(0, height - header_height, width, header_height, fill=True, stroke=False)

        # Add header image if provided
        if header_image_path:
            try:
                header_width = 200
                header_height_image = 100
                can.drawImage(header_image_path, 60, height - header_height_image - 5, width=header_width, height=header_height_image, preserveAspectRatio=True)
            except Exception as e:
                print(f"Error adding header image: {e}")

        # Add the image at placeholder position
        if placeholder_paragraph_index is not None and image_path:
            adjusted_image_path = dpi_adjust_image(image_path)
            if adjusted_image_path:
                img = PILImage.open(adjusted_image_path)
                img.save("temp_image.png")
                target_width_in = 5
                target_height_in = 2
                target_width_pts = target_width_in * 72
                target_height_pts = target_height_in * 72
                y_position = height - (placeholder_paragraph_index + 1) * 100
                can.drawImage("temp_image.png", 100, y_position, width=target_width_pts, height=target_height_pts, preserveAspectRatio=True)

        # Draw footer
        footer_height = 40
        can.setFillColor(light_blue)
        can.rect(0, 0, width, footer_height, fill=True, stroke=False)

        can.setFillColor(colors.black)
        can.setFont("Times-Italic", 8)
        x_position = 50
        y_position = 30
        for line in footnote_text.split("\n"):
            can.drawString(x_position, y_position, line.strip())
            y_position -= 10

        can.save()

        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

# Web Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(docx_path)

        # Convert DOCX to PDF
        try:
            temp_pdf_path = docx_path.replace('.docx', '_temp.pdf')
            pdf_path = convert_docx_to_pdf(docx_path, temp_pdf_path)

            if pdf_path:
                header_image_path = 'static/picture1.png'  # Example header image
                image_path = 'static/Image1.png'  # Example image for <Image1>
                footnote_text = "Global Health Equity. © 2024. Published by Global Health Equity journal."

                # Add image and footnotes to the PDF
                final_pdf_path = docx_path.replace('.docx', '_final.pdf')
                add_images_and_footnote(pdf_path, final_pdf_path, footnote_text, header_image_path=header_image_path, image_path=image_path, docx_path=docx_path)

                # Send the final PDF as a response
                return send_file(final_pdf_path, as_attachment=True)
            else:
                return 'Error in PDF conversion'
        except Exception as e:
            print(f"Error in the upload process: {e}")
            return f"An error occurred: {e}"

    return 'Invalid file format'

if __name__ == '__main__':
    app.run(debug=False)
