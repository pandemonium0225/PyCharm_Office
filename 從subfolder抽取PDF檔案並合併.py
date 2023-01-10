import shutil
import os
import sys
import PyPDF2
from PyPDF2 import PdfFileMerger, PdfFileReader
from natsort_merge import natsort_merge

source = input("請輸入PUBLISHER的發票資料夾位置")


def del_DS_Store():
    for root,dirs,files in os.walk(source,topdown=False):
        merger = PdfFileMerger()
        for file in reversed(files):
            if ".DS_Store" in file:
                del_location = os.path.join(root,file)
                print(del_location)
                os.remove(del_location)
        #     try:
        #         with open(os.path.join(root,file),'rb') as file:
        #             merger.append(PdfFileReader(file),pages=(0,1))
        #     except Exception as e:
        #         print(file + " not correctly merged because...",e)
        # merger.write(root + '.pdf')


def merge_invoice():
    # os.walk 的topdown 選項，為TRUE，則遍歷根目錄，與根目錄中的每個子目錄
    for dirName, subDir,fileList in os.walk(source,topdown='False'):
        merger=PdfFileMerger()
        for fname in reversed(fileList):
            try:
                with open(os.path.join(dirName,fname),'rb') as file:
                    merger.append(PdfFileReader(file),pages=(0,1))
    #             merger.write(dirName + ".pdf")
            except Exception as e:
                print(fname + " not correctly merged, because...",e)
        print(dirName)
        merger.write(dirName + ".pdf")



if __name__ == "__main__":
    del_DS_Store()
    merge_invoice()
    natsort_merge()



