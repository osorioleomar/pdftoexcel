import PyPDF2
import openpyxl
import os
import re

# Function to clean text by removing non-ASCII characters
def clean_text(text):
    # Remove non-ASCII characters
    cleaned_text = ''.join([char if 32 <= ord(char) < 128 else ' ' for char in text])
    
    # Replace spaces with a single space
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    
    return cleaned_text

# Open the PDF file
pdf_file_path = "2013-01-01annual-report-2013-nordea-bank-aben.pdf"
pdf_file = PyPDF2.PdfReader(pdf_file_path)

# Create an Excel workbook
wb = openpyxl.Workbook()
ws = wb.active

# Iterate over the pages in the PDF
for page_num in range(len(pdf_file.pages)):
    # Get the text from the current page
    page_text = pdf_file.pages[page_num].extract_text()
    
    # Clean the text by removing non-ASCII characters and extra spaces
    cleaned_text = clean_text(page_text)
    
    # Write the cleaned text to the Excel workbook
    ws.cell(row=page_num + 1, column=1).value = cleaned_text

# Generate an Excel filename based on the PDF file name
pdf_file_name = os.path.basename(pdf_file_path)
excel_file_name = os.path.splitext(pdf_file_name)[0] + ".xlsx"

# Save the Excel workbook
wb.save(excel_file_name)

print(f"Extracted text from {pdf_file_name} and saved to {excel_file_name}")
