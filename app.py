import re
import openpyxl
from pathlib import Path
import PyPDF2
from docx import Document
import io

def extract_info(cv_file):
  """
  Extracts email IDs, contact numbers, and overall text from a CV based on its format.

  Args:
      cv_file (str): The path to the CV file.

  Returns:
      dict: A dictionary containing email IDs, contact numbers, and overall text.
  """
  email_regex = r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+"
  phone_regex = r"\d{3}-\d{3}-\d{4}|\d{7}|\+\d{1,2}-\d{3}-\d{3}-\d{4}"

  text = ""
  if cv_file.suffix == ".pdf":
    with open(cv_file, 'rb') as pdf_file:
      pdf_reader = PyPDF2.PdfReader(pdf_file)
      for page in pdf_reader.pages:
        text += page.extract_text()
  elif cv_file.suffix in [".doc", ".docx"]:
    with open(cv_file, 'rb') as doc_file:
      if cv_file.suffix == ".docx":
        doc = Document(io.BytesIO(doc_file.read()))
        for paragraph in doc.paragraphs:
          text += paragraph.text
      else:  # Assuming .doc format
        # Implement a library like docx2html or pypiwin32 for doc file processing (more complex)
        # For demonstration purposes, leaving doc format extraction unimplemented.
        print(f"Warning: Doc file format extraction not implemented for: {cv_file}")
  else:
    print(f"Unsupported file format: {cv_file}")
    return None

  emails = re.findall(email_regex, text)
  phones = re.findall(phone_regex, text)
  return {"emails": emails, "phones": phones, "text": text.strip()}

def write_to_excel(data, filename):
  """
  Writes extracted information to an XLSX file.

  Args:
      data (list): A list of dictionaries containing extracted information from each CV.
      filename (str): The filename for the output XLSX file.
  """
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.append(["Emails", "Phones", "Text"])

  for cv_data in data:
    if cv_data:  # Check if data extraction was successful (skip for unsupported formats)
      ws.append([", ".join(cv_data["emails"]), ", ".join(cv_data["phones"]), cv_data["text"]])

  wb.save(filename)

# Define folder path containing CVs
cv_folder_path = Path("Sample2")

# Initialize empty list to store extracted data
extracted_data = []

# Loop through all files in the folder
if __name__ == "__main__":
  for cv_file in cv_folder_path.iterdir():
    if cv_file.is_file():
      extracted_info_dict = extract_info(cv_file)
      if extracted_info_dict:
        extracted_data.append(extracted_info_dict)

  write_to_excel(extracted_data, "extracted_cvs.xlsx")

  print("Information extracted and saved to extracted_cvs.xlsx")
