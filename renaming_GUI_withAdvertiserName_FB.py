from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import os
import re
import pandas as pd
from PDF_Functions import convert_pdf_to_txt
ref_location = r"C:\Users\sebein\Desktop\結帳\FB\2022\Oct\PG_Monthly_Billing_FB.xlsx"
walk = os.walk(r"C:\Users\sebein\Desktop\結帳\FB\2022\Oct\PG_Test")


def get_row(ref_location,job_no):
    job_row = []
    ref_list = pd.read_excel(ref_location)
    for row in range(len(ref_list)):
        if job_no in ref_list['行銷活動名稱'][row]:
            real_no = row + 2
            job_row.append(real_no)
    if job_row:
        return max(job_row)
    else:
        pass


for root, dirs, files in walk:
    for name in files:
        advertiser_list = []
        string = convert_pdf_to_txt(os.path.join(root,name))
        lines = list(filter(bool,string.split('\n')))
        for i in range(len(lines)):
            advertiser_pattern = re.search(r"CP~(PG\d{6})_AN~(\w\s?.*?)_",lines[i])
            if advertiser_pattern is not None:
                job_no = advertiser_pattern.group(1)
                client_name = advertiser_pattern.group(2)
                advertiser_list.append(client_name)
                row = get_row(ref_location,job_no)
            else:
                advertiser_pattern = re.search(r"(SCTWFBA\d{6}).*?_(\w.*?)_",lines[i])
                if advertiser_pattern is not None:
                    job_no = advertiser_pattern.group(1)
                    client_name = advertiser_pattern.group(2)
                    advertiser_list.append(client_name)
                    row = get_row(ref_location,job_no)
        advertiser_set = set(advertiser_list)
        advertiser_unique_list = list(advertiser_set)
        client_name_concat = "@".join(advertiser_unique_list)
        print("Renaming : {} to {}".format(name,"[" + str(row) + "]" + name[:-4] + "_" + client_name_concat + ".pdf"))
        os.rename(os.path.join(root, name), os.path.join(root,"[" + str(row)+ "]" + name[:-4] + "_" + client_name_concat + ".pdf"))



# def add_client_name():
#     walk = os.walk(r"C:\Users\sebein\Desktop\結帳\FB\test")
#     for root, dirs, files in walk:
#         for name in files:
#             advertiser_list = []
#             string = convert_pdf_to_txt(os.path.join(root,name))
#             lines = list(filter(bool,string.split('\n')))
#             for i in range(len(lines)):
#                 advertiser_pattern = re.search(r"CP~(PG\d{6})_AN~(\w\s?.*?)_",lines[i])
#                 if advertiser_pattern is not None:
#                     job_no = advertiser_pattern.group(1)
#                     client_name = advertiser_pattern.group(2)
#                     advertiser_list.append(client_name)
#                 else:
#                     advertiser_pattern = re.search(r"(SCTWFBA\d{6}).*?_(\w.*?)_",lines[i])
#                     if advertiser_pattern is not None:
#                         job_no = advertiser_pattern.group(1)
#                         client_name = advertiser_pattern.group(2)
#                         advertiser_list.append(client_name)
#             advertiser_set = set(advertiser_list)
#             advertiser_unique_list = list(advertiser_set)
#             client_name_concat = "&".join(advertiser_unique_list)
#             os.rename(os.path.join(root, name), os.path.join(root, name[:-4] + "_" + client_name_concat + ".pdf"))
#
# def add_rowno():
#     walk = os.walk(r"C:\Users\sebein\Desktop\結帳\FB\test")
#     for root, dirs, files in walk:
#         for name in files:
#             string = convert_pdf_to_txt(os.path.join(root,name))
#             lines = list(filter(bool,string.split('\n')))
#             for i in range(len(lines)):
#                 if "Account Id / Group:" in lines[i]:
#                     advertiser_id=lines[i+1]
#                     print(advertiser_id)
#                     ref_list=pd.read_excel(r"C:\Users\sebein\Desktop\結帳\FB\2022\Aug\PG_Monthly_Billing_AOD_new_20220826.xlsx")
#                     for n in range(len(ref_list)):
#                         try:
#                             if int(advertiser_id) == int(ref_list['帳號編號'][n]):
#                                 row_no=ref_list[ref_list.帳號編號 == int(advertiser_id)].index.tolist()
#                                 print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name))
#                                 os.rename(os.path.join(root,name),os.path.join(root,"[" + str(row_no[0] + 2)+ "]" + name.replace("/","")))
#                                 break
#                         except PermissionError as p:
#                             print(p)
#                             break
#                         except ValueError as e:
#                             print(e)
#                             break
#                         except FileNotFoundError as f:
#                             print(f)
#                             break
#
#
# if __name__ == "__main__":
#     add_rowno()
#     add_client_name()