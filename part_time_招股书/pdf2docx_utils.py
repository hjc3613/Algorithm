import os

from pdf2docx import Converter

def convert_pdf2docx(pdf, docx):
    # convert pdf to docx
    cv = Converter(pdf)
    cv.convert(docx)      # all pages by default
    cv.close()

def convert_pdf_in_dirs(origin, target):
    for pdf_file in os.listdir(origin):
        pdf_file_path = os.path.join(origin, pdf_file)
        docx_file_path = os.path.join(target, pdf_file.replace('.pdf', '.docx'))
        convert_pdf2docx(pdf_file_path, docx_file_path)

if __name__ == '__main__':
    pdf_file_dir = r'C:\Users\hujunchao\Documents\PdfDir\original\招股书\batch1-split'
    docx_file_dir = r'C:\Users\hujunchao\Documents\PdfDir\original\招股书docx\batch1-次'
    convert_pdf_in_dirs(pdf_file_dir, docx_file_dir)