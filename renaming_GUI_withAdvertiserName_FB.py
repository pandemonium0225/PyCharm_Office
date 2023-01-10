import os
import re
import pandas as pd
from PDF_Functions import convert_pdf_to_txt
ref_location = r"C:\Users\sebein\Desktop\結帳\FB\2022\Dec\PG_Monthly_Billing_AOD_new_20230103.xlsx"
walk = os.walk(r"C:\Users\sebein\Desktop\結帳\FB\2022\Dec\PG")


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
