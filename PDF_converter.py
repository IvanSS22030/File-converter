#First we need to import everything you'll use to make it work with the files of your preference. 
import img2pdf
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from docx import Document
from docx2pdf import convert
from PIL import Image


#next we define a function to convert it to PDF 
def convert_doc_to_pdf(input_path, output_path):
    convert(input_path, output_path)

#Then we do the same, but backwards, PDF to Doc, but we use Page.extractText() to get the text from the file and then we use Doc.save (and then we provide the outpath later in the code)
def convert_pdf_to_doc(input_path, output_path):
    pdf = PdfFileReader(input_path)
    doc = Document()
#This add the new funtionality to convert JPG files to PDF. 
    for page_num in range(pdf.numPages):
        page = pdf.getPage(page_num)
        text = page.extractText()
        doc.add_paragraph(text)

    doc.save(output_path)

def convert_jpg_to_pdf(input_path,output_path):
  Im_1=Image.open(input_path)
  Im_1.save(output_path)  


#We define the function that will make the file convertion 
def convert_file(input_path, output_path):
    file_name, file_extension = os.path.splitext(input_path)
    file_extension = file_extension.lower()

    if file_extension == ".pdf":
        convert_pdf_to_doc(input_path, output_path)
    if file_extension ==".jpg":
        convert_jpg_to_pdf(input_path, output_path)   
    elif file_extension == ".docx":
        convert_doc_to_pdf(input_path, output_path)
    else:
        print("Unsupported file format.")

#We specify the routes, one inputh and the output, being careful of the / slashes. 
input_file_path = "C:/Users/IvanJ/Downloads/Bankstatement.jpg"
output_file_path = "C:/Users/Ivanj/OneDrive/Escritorio/foto.pdf"
convert_file(input_file_path, output_file_path)
