import os
from fnmatch import fnmatch
from PyPDF2 import PdfReader, PdfWriter

root = './'
pattern = "*.pdf"
pdffiles = [];

def merge_pdfs(paths, output):
    pdf_writer = PdfWriter()

    for path in paths:
        pdf_reader = PdfReader(path)
        for page in range(len(pdf_reader.pages)):
            # Add each page to the writer object
            pdf_writer.add_page(pdf_reader.pages[page])

    # Write out the merged PDF
    with open(output, 'wb') as out:
        pdf_writer.write(out)

def main():
    for rut in os.listdir(root):
        if(os.path.isdir(rut)):
            for path, subdirs, files in os.walk(rut):
                for name in files:
                    if fnmatch(name, pattern):
                        pdffiles.append(os.path.join(path, name))
                #print(path)
            merge_pdfs(pdffiles, output=rut + '/' + rut + '.penalolen.pdf')    
            pdffiles.clear()

        
if __name__ == "__main__":
    main()
   