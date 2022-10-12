from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import os
import pandas as pd
from PDF_Functions import convert_pdf_to_txt


def rename_PG_FB():
    file_loc = input("pls input the directory you want to rename!!!")
    ref_loc = input("pls input the ref file location with file name!!!")
    ref_list = pd.read_excel(ref_loc)
    walk = os.walk(file_loc)
    for root, dirs, files in walk:
        for name in files:
            string = convert_pdf_to_txt(os.path.join(root,name))
            lines = list(filter(bool,string.split('\n')))
            for i in range(len(lines)):
                if "Account Id / Group:" in lines[i]:
                    advertiser_id=lines[i+1]
                    for n in range(len(ref_list)):
                        try:
                            if int(advertiser_id) == int(ref_list['帳號編號'][n]):
                                row_no=ref_list[ref_list.帳號編號 == int(advertiser_id)].index.tolist()
                                print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name))
                                os.rename(os.path.join(root,name),os.path.join(root,"[" + str(row_no[0] + 2)+ "]" + name.replace("/","")))
                                break
                        except PermissionError as p:
                            pass
                            # print(p)
                        except ValueError as e:
                            pass
                            # print(e)
                        except FileNotFoundError as f:
                            pass
                            # print(f)

if __name__ == "__main__":
    rename_PG_FB()