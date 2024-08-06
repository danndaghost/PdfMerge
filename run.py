import os
from openpyxl import load_workbook
from fnmatch import fnmatch
from PyPDF2 import PdfReader, PdfWriter

#rootpath = './'
pattern = "*.pdf"
rootpath = '\\\\192.168.0.11\\cormup\\Dir_Personas\\Contratos-General'
writepath = "Contratos SLEP"
readpath = ['PC Bastian', 'Publicacion']
excludepath = ['NO VIGENTE']
pdffiles = []
dotacion = []

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

def leerExcel():
    filepath        = rootpath
    archivo         = filepath + '\\' + writepath + '\\' + 'dotacion'    
    archivo_excel   = archivo + '.xlsx'
    workbook = load_workbook(archivo_excel)
    hoja = workbook.worksheets[0]
    for row in range(2, hoja.max_row+1): #start on two to ignore the title of column
        dotacion.append(hoja.cell(row=row, column=1).value)
    return dotacion

def crearPdf(readpath, dotacion):
    contador = 0
    cantidadtotal = len(dotacion)
    for rut in dotacion:
        contador = contador+1
        print(":::::::::::::::  LEYENDO  :::  " + str(contador) + "  de  " + str(cantidadtotal) +  "  ::::  RUT  " + rut + "    :::::::::::::::::::")
        for rootfolder in readpath:
            for path, dirs, files in os.walk(rootpath + '\\' + rootfolder):
                dirs[:] = [d for d in dirs if d not in excludepath]
                for name in files:
                    if fnmatch(name, pattern) and (rut in name):
                        print(os.path.join(path, name))
                        pdffiles.append(os.path.join(path, name))
        
        if len(pdffiles) > 0:
            merge_pdfs(pdffiles, output=rootpath + '\\' + writepath + '\\' + rut + '.penalolen.pdf')    
        pdffiles.clear()
        
if __name__ == "__main__":
    dotacion = leerExcel() #obtener los rut a buscar en las carpetas
    crearPdf(readpath, dotacion)
   