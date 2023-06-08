from docx2pdf import convert
from PyPDF2 import PdfMerger
from os import listdir, mkdir
from os.path import isfile, join, exists
import sys

def docx2pdf(path):
    if not exists(path+'\\pdf'):
         mkdir(path+'\\pdf')
    convert(path, path+'\\pdf')
    print("\nDocx to Pdf conversion completed successfully")
    
def pdf_merger(path):
    path += '\\pdf'
    pdf_files = [f for f in listdir(path) if isfile(join(path,f)) and '.pdf' in f]
    print("\nList of files:")
    for f in pdf_files:
         print(f)
    resultFile = input("\nInsert the name of the result file: ")
    if '.pdf' not in resultFile:
         resultFile += '.pdf'
    
    merger = PdfMerger()
    for pdf in pdf_files:
         merger.append(path+'\\'+pdf)
    
    if not exists(path+'\\Output'):
         mkdir(path+'\\Output')
    merger.write(path+'\\Output\\'+resultFile)
    merger.close()
    print("\nPdf files merged successufully")


if __name__ == "__main__":
     path = input("\nInserisci il path dei file docx")
     docx2pdf(path)
     pdf_merger(path)



