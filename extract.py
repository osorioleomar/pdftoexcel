import os
import PyPDF2
import openpyxl

def extract_text_from_pdf(pdf_file):
  """Extracts all the text from a PDF file.

  Args:
    pdf_file: The path to the PDF file.

  Returns:
    A string containing all the text in the PDF file.
  """

  pdf_reader = PyPDF2.PdfReader(pdf_file)
  text = ""
  for page in pdf_reader.pages:
    text += page.extract_text()

  # Replace all occurrences of the character "◦" with an empty string.
  text = text.translate(str.maketrans("", "", "◦"))

  return text

def create_excel_file(excel_file_path):
  """Creates an Excel file with a single worksheet.

  Args:
    excel_file_path: The path to the Excel file.
  """

  workbook = openpyxl.Workbook()
  worksheet = workbook.active
  worksheet.title = "Text"
  workbook.save(excel_file_path)

def write_text_to_excel_file(excel_file_path, text):
  """Writes text to an Excel file.

  Args:
    excel_file_path: The path to the Excel file.
    text: The text to write to the Excel file.
  """

  workbook = openpyxl.load_workbook(excel_file_path)
  worksheet = workbook.active
  worksheet.append(text.split("\n"))
  workbook.save(excel_file_path)

if __name__ == "__main__":
  pdf_folder_path = "PDF"
  pdf_files = os.listdir(pdf_folder_path)

  for pdf_file in pdf_files:
    # Extract the text from the PDF file.
    text = extract_text_from_pdf(os.path.join(pdf_folder_path, pdf_file))

    # Create an Excel file with the same filename as the PDF file.
    excel_file_path = os.path.join(pdf_folder_path, pdf_file.replace(".pdf", ".xlsx"))
    create_excel_file(excel_file_path)

    # Write the text to the Excel file.
    write_text_to_excel_file(excel_file_path, text)

    print("Extracted the text from {} to {}".format(pdf_file, excel_file_path))