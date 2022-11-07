"""
若下載新的DV360 REPORT，可用此檢查與舊的REPORT的差異及更新REPORT
"""
import os
import pandas as pd
import warnings
warnings.simplefilter(action='ignore',category=UserWarning)
df1 = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Oct\Monthly_File_202210.xlsx")
df2 = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Oct\Monthly_Accrual_Spend_File_20221031.xlsx")
root = r"C:\Users\sebein\Desktop\結帳\DBM\2022\Oct"
filename = "updated_dv360_file.xlsx"
updated_file_location = os.path.join(root,filename)

def check_difference():
    # df1 = pd.read_excel(input("請輸入之前下載的DV360_report檔案位置"))
    # df2 = pd.read_excel(input("請輸入新下載的DV360_report檔案位置"))
    oldnotinnew = df2[~(df2['Insertion Order ID'].isin(df1['Insertion Order ID']))].dropna().reset_index(drop=True)
    df_diff = pd.concat([df1, df2]).drop_duplicates(['Insertion Order ID'],keep=False)
    if df_diff.empty:
        print("no updated campaigns found")
    else:
        df_diff.to_excel(updated_file_location)
    return oldnotinnew, df_diff


if __name__ == "__main__":
    check_difference()