from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import re
import pandas as pd
import sys, os, PyPDF2
from natsort import natsorted, ns
from Extract_Subfolders import extract_sub
from Extract_DBM_Amount import dbm_amount

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos= 0
    pageno=1
    fstr = ''
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
        str = retstr.getvalue()
        fstr += str

    fp.close()
    device.close()
    retstr.close()
    return fstr

def rename_invoice():
    walk_directory = input("please input the directory within it you want to rename")
    walk=os.walk(walk_directory)
    for root, dirs, files in walk:
        for name in files:
            print(name)
            string = convert_pdf_to_txt(os.path.join(root,name))
            lines = list(filter(bool,string.split('\n')))
            for i in range(len(lines)):
                if "Advertiser Id:" in lines[i]:
                    advertiser_id=lines[i].split(':',1)
                    advertiser_id2=re.search(r"\d{7,9}",advertiser_id[1])
                    print(advertiser_id2.group(),advertiser_id)
                    ref_list=pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Mar\Monthly_File_20220329_test.xlsx")
                    for n in range(len(ref_list)):
                        try:
                            if int(advertiser_id2.group()) == int(ref_list['AdvertiserID'][n]):
                                row_no=ref_list[ref_list.AdvertiserID == int(advertiser_id2.group())].index.tolist()
                                print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name))
                                os.rename(os.path.join(root,name),os.path.join(root,"[" + str(row_no[0] + 2)+ "]" + name.replace("/","")+".pdf"))
                                break
                        except:
                            continue
                    break

def sort_invoice_file():
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


if __name__ == '__main__':
    # extract_sub()
    rename_invoice()
    dbm_amount()
    sort_invoice_file()
