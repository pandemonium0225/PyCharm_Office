import sys, os, PyPDF2
from natsort import natsorted, ns


def natsort_merge():
    userpdflocation = input("Folder path to PDFs that need merging:")
    # os.chdir(userpdflocation)
    userfilename = input("What should the file be called?")
    pdf2merge = []
    for filename in os.listdir(userpdflocation):
        if filename.endswith('.pdf'):
            pdf2merge.append(filename)

    pdf_nature_merge = natsorted(pdf2merge, key=lambda y: y.lower())

    pdfWriter = PyPDF2.PdfFileWriter()

    for filename in pdf_nature_merge:
        with open(userpdflocation + '\\' + filename, 'rb') as pdfFileObj:
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            for pageNum in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)

            with open(userpdflocation + '\\' + userfilename + '.pdf', 'wb') as pdfOutput:
                pdfWriter.write(pdfOutput)
    print("The designated files has been merged and you shall find it in ", str(userpdflocation))

if __name__ == "__main__":
    natsort_merge()