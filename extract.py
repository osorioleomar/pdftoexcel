import PyPDF2
import openpyxl

# Create an Excel workbook
wb = openpyxl.Workbook()

# Open the PDF file
pdf_file = PyPDF2.PdfReader("csr2017e_print.pdf")

# Iterate over the pages in the PDF
for page_num in range(len(pdf_file.pages)):

    # Get the text from the current page
    page_text = pdf_file.pages[page_num].extract_text()

    # Write the text to the Excel workbook
    ws = wb.active
    ws.cell(row=page_num + 1, column=1).value = page_text

# Save the Excel workbook
wb.save("csr2017e_print.xlsx")