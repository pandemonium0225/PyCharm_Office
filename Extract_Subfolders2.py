import sys,os,PyPDF2
from PyPDF2 import PdfFileMerger,PdfFileReader
from natsort import natsorted,ns
# pass the path of the parent_folder
def fetch_all_files(parent_folder: str):
    target_files = []
    for path, subdirs, files in os.walk(parent_folder,topdown='True'):
        for name in files:
            target_files.append(os.path.join(path, name))
    print(target_files)
    return target_files


# pass the path of the output final file.pdf and the list of paths
def merge_pdf(out_path, extracted_files):
    merger = PdfFileMerger()

    for pdf in reversed(extracted_files):
        merger.append(pdf)

    merger.write(out_path)
    merger.close()

def nat_merge():
    userpdflocation = input("Folder path to PDFs that need merging:")
    os.chdir(userpdflocation)
    userfilename = input("What should the file be called?")
    pdf2merge = []
    for filename in os.listdir('..'):
        if filename.endswith('.pdf'):
            pdf2merge.append(filename)

    pdf_nature_merge = natsorted(pdf2merge, key=lambda y: y.lower())

    pdfWriter = PyPDF2.PdfFileWriter()

    for filename in pdf_nature_merge:
        with open(filename, 'rb') as pdfFileObj:
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            for pageNum in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)

            with open(userfilename + '.pdf', 'wb') as pdfOutput:
                pdfWriter.write(pdfOutput)
    print("Done")



# get a list of all the paths of the pdf
# parent_folder_path = r'C:\Users\sebein\Desktop\結帳\DBM\2022\Mar\2022_03'
# outup_pdf_path = r'C:\Users\sebein\Desktop\結帳\DBM\2022\Mar\2022_03\invoices.pdf'
#
# extracted_files = fetch_all_files(parent_folder_path)
# merge_pdf(outup_pdf_path, extracted_files)

if __name__ == "__main__":
    all_file = fetch_all_files(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Merge_test")
    merge_pdf(r"D:\Merge_test", all_file)
