import os.path
import traceback

from tqdm import tqdm

from PyPDF2 import PdfFileWriter, PdfFileReader

def pdf_spliter(in_file, out_prefix, max_pages):
    assert out_prefix.endswith('.pdf')
    inputpdf = PdfFileReader(open(in_file, "rb"))
    if inputpdf.is_encrypted:
        print('加密文件：', in_file)
        return
    total_pages = inputpdf.getNumPages()
    sub_pages = [(i, min(i+max_pages, total_pages)) for i in range(0, total_pages, max_pages)]
    for idx, (start, end) in enumerate(sub_pages):
        try:
            output_path = out_prefix.replace('.pdf', f'-{idx}-{start}-{end}.pdf')
            output = PdfFileWriter()
            for i in range(start, end):
                output.add_page(inputpdf.getPage(i))
            with open(output_path, "wb") as outputStream:
                output.write(outputStream)
        except:
            print('error: ', in_file, idx, start, end)
            traceback.format_exc()

def main():
    root = r'C:\Users\hujunchao\Documents\PdfDir\original\招股书'
    sub_dir_in = [ 'batch10']
    sub_dir_out = [ 'batch10-split']
    for dir_in, dir_out in zip(sub_dir_in, sub_dir_out):
        assert dir_out.split('-')[0] == dir_in
        dir_in = os.path.join(root, dir_in)
        dir_out = os.path.join(root, dir_out)
        for file in tqdm(os.listdir(dir_in)):
            path_in = os.path.join(dir_in, file)
            path_out = os.path.join(dir_out, file)
            pdf_spliter(path_in, path_out, max_pages=100)
if __name__ == '__main__':
    main()