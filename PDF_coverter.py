import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from docx import Document
from docx2pdf import convert


def convert_doc_to_pdf(input_path, output_path):
    convert(input_path, output_path)


def convert_pdf_to_doc(input_path, output_path):
    pdf = PdfFileReader(input_path)
    doc = Document()

    for page_num in range(pdf.numPages):
        page = pdf.getPage(page_num)
        text = page.extractText()
        doc.add_paragraph(text)

    doc.save(output_path)


def convert_file(input_path, output_path):
    file_name, file_extension = os.path.splitext(input_path)
    file_extension = file_extension.lower()

    if file_extension == ".pdf":
        convert_pdf_to_doc(input_path, output_path)
    elif file_extension == ".docx":
        convert_doc_to_pdf(input_path, output_path)
    else:
        print("Unsupported file format.")


input_file_path = "C:/Users/IvanJ/OneDrive/Escritorio/My python journey.docx"
output_file_path = "C:/Users/Ivanj/OneDrive/Escritorio/test.pdf"
convert_file(input_file_path, output_file_path)
