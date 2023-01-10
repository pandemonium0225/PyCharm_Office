"""
確認條件:
2. 當月開立發票的當月MEDIA SCHEDULE P2 SUBTOTAL金額加上跨月部分要等於BRIEF金額
"""

import pandas as pd
import os
import openpyxl
import numpy as np


def check_budget():
    for root,dirs,filelist in os.walk(r"C:\Users\sebein\Desktop\makeup schedule\CheckMediaScheduleNumber"):
        for file in filelist:
            if file.endswith('xlsx'):
                file_dir = os.path.join(root,file)
                df = pd.read_excel(file_dir,'Daily - Broadcast',skiprows=16,usecols='T,U,W',nrows=10)
                df.columns = ['NetCost','InsertionNumber','TotalCostofMedia']
                df = df[(df[['NetCost','InsertionNumber','TotalCostofMedia']] !=0).any(axis=1)]
                workbook = openpyxl.load_workbook(file_dir)
                worksheet = workbook['Daily - Broadcast']
                this_month_amount = worksheet['W62'].value
                schedule_name = worksheet['D10'].value
                billing_entity = worksheet['D8'].value
                brand = worksheet['D9'].value
                print(schedule_name)
        #         print(df)
                if 'TW1713' in billing_entity:
                    print(brand, "_",schedule_name, "PFX本月將開立發票給APEX", round(this_month_amount,0), "元")
                    print('*' * 100)
                elif len(df[df['InsertionNumber'] ==0]) != 0:
                    if ('TW9098') or ('TW8081') in billing_entity:
                        print(brand, "_",schedule_name, "本月將開立發票", round(this_month_amount + previous_amount,0), "元")
                        print('*' * 100)
                    else:
                        try:
                            print(brand, "_",schedule_name, "本月將開立發票", round(this_month_amount,0), "元")
                            print('*'*100)
                        except Exception as e:
                            print(schedule_name + "not working because of", e)
                elif len(df[df['InsertionNumber'] ==0]) == 0:
                    if 'TW8081' or 'TW9098' in billing_entity:
                        print(brand, "_",schedule_name, "本月將開立發票", round(this_month_amount,0), "元")
                        print('*' * 100)

                    elif 'TW1713' in billing_entity:
                        print(brand, "_",schedule_name, "PFX本月將開立發票給APEX", round(this_month_amount,0), "元")
                        print('*' * 100)

def check_p1p2_diff():
    count = 0
    p1_dict = {}
    p2_dict = {}
    for root, dirs, filelist in os.walk(r"C:\Users\sebein\Desktop\makeup schedule\Lay's_FB"):
        for file in filelist:
            if file.endswith('.xlsx'):
                file_dir = os.path.join(root, file)
                df = pd.read_excel(file_dir, 'Daily - Broadcast', skiprows=16, usecols='T,U,W', nrows=10)
                df.columns = ['NetCost', 'InsertionNumber', 'TotalCostofMedia']
                df = df[(df[['NetCost', 'InsertionNumber', 'TotalCostofMedia']] != 0).any(axis=1)]
                workbook = openpyxl.load_workbook(file_dir)
                worksheet = workbook['Daily - Broadcast']
                sub_total = worksheet['W60'].value
                this_month_amount = worksheet['W62'].value
                schedule_name = worksheet['D10'].value
                billing_entity = worksheet['D8'].value
                brand = worksheet['D9'].value
                if 'TW1713' in billing_entity:
                    p1_dict[schedule_name] = sub_total
                elif 'TW9098' or 'TW8081' in billing_entity:
                    p2_dict[schedule_name] = sub_total
        #         print(schedule_name, sub_total,billing_entity)

    shared_items = {k: p1_dict[k] for k in p1_dict if k in p2_dict and p1_dict[k] != p2_dict[k]}
    if len(shared_items) == 0:
        print('All Schedules with the same amount checked!!!')
    else:
        print(shared_items, "with inconsistent amount!!!")


if __name__ == "__main__":
    check_p1p2_diff()