import sys,os,PyPDF2
from PyPDF2 import PdfFileMerger,PdfFileReader


def extract_sub():
    parent_folder = input("please enter the directory you want to extract")
    print(parent_folder)
    for subfolder in os.listdir(parent_folder):
        merger = PdfFileMerger()
        if os.path.isdir(parent_folder + subfolder + '/'):
            for file in reversed(os.listdir(parent_folder + subfolder + '/')):
                if file.endswith('.pdf'):
                    print(file)
                merger.append(parent_folder + subfolder + '\\' + file)
            merger.write(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Merge_test\\" + file)
            merger.close()

def merge_and_extract():
    """抽取子資料夾中的檔案，合併後另存至上一層資料夾"""
    rootDir = input("please input the directory you want to merge and extract!!!")
    for dirName, subDir,fileList in os.walk(rootDir, topdown='True'):
        merger = PdfFileMerger()
        try:
            for fname in reversed(fileList):
                merger.append(PdfFileReader(open(os.path.join(dirName,fname),'rb')))
            merger.write(str(dirName) + ".pdf")
        except Exception as e:
            print(fname + " not correctly merged, because...", e)
            continue
        merger.close()


def save_to_invoice_folder():
    rootDir = r"C:\Users\sebein\Desktop\結帳\DBM\2022\Sep\2022-09"
    files = os.listdir(rootDir)
    os.makedirs(rootDir + "\\invoices")
    save_location = rootDir + "\\invoices\\"
    for f in files:
        fullpath = os.path.join(rootDir,f)
        if os.path.isfile(fullpath):
            os.replace(fullpath,os.path.join(save_location+f))





if __name__ == "__main__":
    merge_and_extract()
    save_to_invoice_folder()