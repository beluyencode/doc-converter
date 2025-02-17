from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter
from docx import Document
from docx.shared import Inches

# Create a new Word document
document = Document()

# Open a PDF file
with open("test.pdf", "rb") as file:
    
    # Create a PdfReader object
    pdf_reader = PdfReader(file)

    # Open the Word document for writing
    with open("output.docx", "wb") as output_file:

        # Loop through each page of the PDF file
        for page_num in range(len(pdf_reader.pages)):
            
            # Get the current page
            page = pdf_reader.pages[page_num]

            # Extract text from the page
            text = page.extract_text()
            
            # Add a paragraph in Word to hold the text
            document.add_paragraph(text)

# Save the Word document
document.save("output.docx")