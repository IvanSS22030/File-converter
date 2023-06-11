import os
import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert
from PIL import Image
import win32com.client

def convert_doc_to_pdf(input_path, output_path):
    convert(input_path, output_path)

def convert_jpg_to_pdf(input_path, output_path):
    Im_1 = Image.open(input_path)
    Im_1.save(output_path, "PDF", quality=100)

def convert_file():
    input_file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("JPEG Files", "*.jpg"), ("Word Document", "*.docx")])
    if input_file_path:
        file_name, file_extension = os.path.splitext(input_file_path)
        file_extension = file_extension.lower()
        if file_extension == ".pdf":
            output_file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
            if output_file_path:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = 0
                wb = word.Documents.Open(input_file_path)
                wb.SaveAs2(output_file_path, FileFormat=16)
                wb.Close()
                print("PDF to DOCX conversion is completed.")
        elif file_extension == ".jpg":
            output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if output_file_path:
                convert_jpg_to_pdf(input_file_path, output_file_path)
                print("JPG to PDF conversion is completed.")
        elif file_extension == ".docx":
            output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if output_file_path:
                convert_doc_to_pdf(input_file_path, output_file_path)
                print("DOCX to PDF conversion is completed.")
        else:
            print("Unsupported file format.")
    else:
        print("No file selected.")

def create_gui():
    root = tk.Tk()
    root.title("File Converter")
    root.geometry("400x200")  # Set the window size (width x height)

    browse_button = tk.Button(root, text="Seleccione su archivo", command=convert_file)
    browse_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
