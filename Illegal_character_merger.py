import sys,os,PyPDF2
from PyPDF2 import PdfFileMerger,PdfFileReader
rootDir=r"C:\Users\sebein\Desktop\結帳\DBM\2022\Apr\2022-04"
for dirName, subDir,fileList in os.walk(rootDir,topdown='True'):
    merger=PdfFileMerger()
    for fname in reversed(fileList):
        merger.append(PdfFileReader(open(os.path.join(dirName,fname),'rb')))
    merger.write(str(dirName)+".pdf")
