from pptx import Presentation
from pptx.util import Inches
import os
import pdfplumber
from docx import Document
import openpyxl

def extract_text_from_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def extract_text_from_pdf(file_path):
    text = ""
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_xlsx(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    content = ""
    for row in sheet.iter_rows(values_only=True):
        content += " | ".join(map(str, row)) + "\n"
    return content

def create_pptx(content, output_file):
    prs = Presentation()

    slides = content.strip().split("\n\n")
    for slide_content in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title, body = slide_content.split("\n", 1) if "\n" in slide_content else (slide_content, "")
        
        slide.shapes.title.text = title.strip()
        slide.placeholders[1].text = body.strip()

    prs.save(output_file)
    print(f"PPTX file created successfully: {output_file}")

def process_file(file_path):
    extension = os.path.splitext(file_path)[-1].lower()
    if extension == '.txt':
        return extract_text_from_txt(file_path)
    elif extension == '.pdf':
        return extract_text_from_pdf(file_path)
    elif extension == '.docx':
        return extract_text_from_docx(file_path)
    elif extension == '.xlsx':
        return extract_text_from_xlsx(file_path)
    else:
        print(f"Unsupported file type: {extension}")
        return ""

def main():
    input_file = input("Enter the path of the document: ").strip()
    content = process_file(input_file)

    if content:
        create_pptx(content, "GeneratedPresentation.pptx")
    else:
        print("No content extracted or unsupported file type.")

if __name__ == "__main__":
    main()
