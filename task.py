import os
import re
import pandas as pd
import fitz 
from pdfplumber import open as open_pdf
from docx import Document

def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    paragraphs = [p.text for p in doc.paragraphs]
    return "\n".join(paragraphs)

def extract_contacts(text):
    # Email regex
    email_regex = r'[\w\.-]+@[\w\.-]+'
    emails = re.findall(email_regex, text)

    # Phone number regex
    phone_regex = r'\b\d{10}\b'
    phones = re.findall(phone_regex, text)

    return emails, phones

def process_cv(cv_dir):
    cv_data = []
    for file_name in os.listdir(cv_dir):
        file_path = os.path.join(cv_dir, file_name)
        if file_name.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif file_name.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        else:
            continue

        emails, phones = extract_contacts(text)
        cv_data.append({'File Name': file_name, 'Email': emails, 'Phone': phones, 'Text': text})

    return cv_data

def create_excel(cv_data, output_file):
    df = pd.DataFrame(cv_data)
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    # Directory containing CVs
    cv_directory = r"C:\Users\dell\Downloads\Sample2-20240406T093029Z-001\Sample2"

    # Process CVs
    cv_data = process_cv(cv_directory)

    # Output Excel file
    output_file = "cv_data.xlsx"

    # Create Excel
    create_excel(cv_data, output_file)

    print("Excel file created successfully.")
