import os
import pandas as pd
from tracker_info import *

def update_dv360_report():
    original_df = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\Monthly_File_20221227202212.xlsx")
    # print(original_df.head())
    update_df = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\Monthly_Accrual_Spend_File_20221229.xlsx",skipfooter=15)
    update_df.rename(columns={"Advertiser ID":"AdvertiserID"},inplace=True)
    update_df['AdvertiserID'] = update_df['AdvertiserID'].astype('str')

    Merged_df = (pd.concat([original_df,update_df],ignore_index=True,sort=False).drop_duplicates(['Insertion Order ID','Insertion Order'],keep='first'))

    Merged_df.to_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\\" + "Monthly_File_20221229_Merged_" + ".xlsx", index=False)
    return Merged_df

if __name__ == '__main__':
    df_monthly = update_dv360_report()
    df_monthly = dv360_spending_file()
    df_monthly = extract_job(df_monthly)
    op_normal(df_monthly)
    pmp_apex(df_monthly)
    op_apex(df_monthly)
    pg_dv360(df_monthly)
    df_monthly.to_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\\" + "Monthly_File_20221229_Merged" + postdate + ".xlsx",
                        index=False)