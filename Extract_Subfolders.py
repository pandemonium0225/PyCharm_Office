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

if __name__ == "__main__":
    extract_sub()