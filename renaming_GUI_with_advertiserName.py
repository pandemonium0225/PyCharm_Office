import os
import re
import pandas as pd
from PDF_Functions import convert_pdf_to_txt
def renaming_with_advertiser_DV360():
    gui_dir = input("Please input the gui directory under which you want to rename!!!")
    walk = os.walk(gui_dir)
    for root, dirs, files in walk:
        for name in files:
            string = convert_pdf_to_txt(os.path.join(root,name))
            lines = list(filter(bool, string.split('\n')))
            for i in range(len(lines)):
                # "Advertiser Id" 欄位需自行轉換為數字，或在擷取TRACKER DATA時一併處理 2022/09/02 仍然需要手動調整AdvertiserID型別
                if "Advertiser Id:" in lines[i]:
                    advertiser_id = lines[i].split(':',1)
                    advertiser_id2 = re.search(r"\d{7,9}", advertiser_id[1])
                    ref_list = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Nov\Monthly_File_added202211.xlsx")
                    for n in range(len(ref_list)):
                        try:
                            if int(advertiser_id2.group()) == int(ref_list['AdvertiserID'][n]):
                                row_no = ref_list[ref_list.AdvertiserID == int(advertiser_id2.group())].index.tolist()
                                advertiser = ref_list['Advertiser'][row_no[0]]
                                advertiser_pattern = re.search(r"AN~(\w+\s?.*?)_MK~TW",advertiser)
                                if advertiser_pattern is None:
                                    print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name[:-4] + "_" + advertiser))
                                    os.rename(os.path.join(root, name), os.path.join(root,"[" + str(row_no[0] + 2) + "]" + name[:-4] + "_" + advertiser + ".pdf"))
                                    break
                                else:
                                    print("Renaming : {} to {}".format(name, "[" + str(row_no[0] + 2)+ "]" + name[:-4] + "_" + advertiser_pattern.group(1)))
                                    os.rename(os.path.join(root, name), os.path.join(root, "[" + str(row_no[0] + 2) + "]" + name[:-4] + "_" + advertiser_pattern.group(1) + ".pdf"))
                                    break
                        except Exception as e:
                            pass
                    break

ref_location = r"C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\Monthly_File_20221227202212.xlsx"
walk = os.walk(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\2022-12")
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

def renaming_with_advertiser_FB():
    ref_location = r"C:\Users\sebein\Desktop\結帳\FB\2022\Nov\PG_Monthly_Billing_AOD_new_20221205.xlsx"
    walk = os.walk(r"C:\Users\sebein\Desktop\結帳\FB\2022\Nov\PG")
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

if __name__ == "__main__":
    try:
        platform_choice = int(input("please input the platform choice 1.DV360 2.FB..."))
        if platform_choice == 1:
            renaming_with_advertiser_DV360()
        elif platform_choice == 2:
            renaming_with_advertiser_FB()
        else:
            print("Please input the valid choice...")
    except Exception as e:
        print("You need to enter the valid option for this process!!!")