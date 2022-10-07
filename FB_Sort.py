import os
import pandas as pd
from renaming_dv360 import convert_pdf_to_txt
from renaming_dv360 import sort_invoice_file

def fb_sort_el():
    walk=os.walk(r"C:\Users\sebein\Desktop\結帳\FB\2022\Sep\PG\PDF")
    for root, dirs, files in walk:
        for name in files:
            string = convert_pdf_to_txt(os.path.join(root,name))
            lines = list(filter(bool,string.split('\n')))
            for i in range(len(lines)):
                if "Account Id / Group:" in lines[i]:
                    advertiser_id=lines[i+1]
                    # print(advertiser_id)
                    ref_list=pd.read_excel(r"C:\Users\sebein\Desktop\結帳\FB\2022\Sep\PG_Monthly_Billing_AOD_202209.xlsx")
                    for n in range(len(ref_list)):
                        try:
                            if int(advertiser_id) == int(ref_list['Account_ID'][n]):
                                row_no=ref_list[ref_list.Account_ID == int(advertiser_id)].index.tolist()
                                print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name))
                                os.rename(os.path.join(root,name),os.path.join(root,"[" + str(row_no[0] + 2)+ "]" + name.replace("/","")+".pdf"))
                                continue
                        except PermissionError:
                            pass
                        except ValueError as e:
                            pass
                        except FileNotFoundError:
                            pass

def fb_sort_pg():
    pg_folder = os.walk(input("please input the pg directory you want to sort"))
    ref_list = pd.read_excel(input("Please enter the ref file location"))
    for root, dirs, files in pg_folder:
        for name in files:
            string = convert_pdf_to_txt(os.path.join(root, name))
            lines = list(filter(bool, string.split('\n')))
            for i in range(len(lines)):
                if "Account Id / Group:" in lines[i]:
                    advertiser_id = lines[i + 1]
                    print(advertiser_id)
                    for n in range(len(ref_list)):
                        try:
                            if int(advertiser_id) == int(ref_list['帳號編號'][n]):
                                row_no=ref_list[ref_list.帳號編號 == int(advertiser_id)].index.tolist()
                                print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name))
                                os.rename(os.path.join(root,name),os.path.join(root,"[" + str(row_no[0] + 2)+ "]" + name.replace("/","")+".pdf"))
                                break
                        except PermissionError as p:
                            print(p)
                        except ValueError as e:
                            print(e)
                        except FileNotFoundError as f:
                            print(f)

                        if str(advertiser_id) == str(ref_list['帳號編號'][n]):
                            print(advertiser_id,ref_list['帳號編號'][n])
                        try:
                            if int(advertiser_id) == int(ref_list['帳號編號'][n]):
                                row_no=ref_list[ref_list.帳號編號 == int(advertiser_id)].index.tolist()
                                print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name))
                                os.rename(os.path.join(root,name),os.path.join(root,"[" + str(row_no[0] + 2)+ "]" + name.replace("/","")+".pdf"))
                            else:
                                pass
                        except Exception as e:
                            print(e)
                            continue
                    break

if __name__ == '__main__':
    # fb_sort_el()
    fb_sort_pg()