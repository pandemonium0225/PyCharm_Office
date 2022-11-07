import pygsheets
import pandas as pd
import re
import datetime
from time import time

"""
1. 從GoogleAds系統後台下載當月報表，內含["Account name","Customer ID", "Budget name", "Currency code", "Billed cost"]等欄位
"""

today = datetime.datetime.now()
postdate = (str(today.year) + str(today.month).zfill(2))
gc = pygsheets.authorize(service_file=r'C:\Users\sebein\Desktop\Pythonupload\pythonupload-307303-aae13f0bea1f.json')
sht_op = gc.open_by_url(
    'https://docs.google.com/spreadsheets/d/12Ts2BP9acoemXb-0bLgdBl8G6PQyMUEoUXrB0b3W8Fo/edit#gid=1390695304'
)
sht_normal = gc.open_by_url(
    'https://docs.google.com/spreadsheets/d/1J_6W-5VE55M4-hjvP0m-hwU7RFRYgRmhVehS_LG5UWM/edit#gid=1014058029'
)
wks = sht_op.worksheets()
wks_op = wks[1]
df_op_Apex = wks_op.get_as_df(start="A1", numerize=False, include_tailing_empty=False)
wks_normal = sht_normal.worksheets()
df_normal = wks_normal[0].get_as_df(start="A1", numerize=False, include_tailing_empty=False)

# 將下載報表轉換為dataframe，目前讀檔預設為excel檔
GoogleAds_monthly = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\Adwords\2022\2022_10\Billed cost_20221104.xlsx",
                                  skiprows=2,
                                  index_col=None,
                                  na_values=['NA']
                                  )
# 下載報表的DATAFRAME中加入欄位
GoogleAds_monthly = GoogleAds_monthly.assign(start='', end='', budget='', owner='', jobnumber='')

# 建立REGEX規則
regex_list = [
    r"(?:PFTW.*?\d{6}.*?])",
    r"(?:APEX|PMSC|PMZN|PMTW)\d{6}_.*?_",
    r"(?:ZNTWGDV|SCTWGDV|PMTWGDV|SCTWGAD|ZNTWGAD)\d{6}.*?"
]
# 抽取JOB NUMBER/剔除非博丰命名的CAMPAIGN/重新設定INDEX
GoogleAds_monthly['jobnumber'] = GoogleAds_monthly['Budget name'].str.extract("(" + "|".join(regex_list) + ")")
GoogleAds_monthly = GoogleAds_monthly[GoogleAds_monthly['jobnumber'].notna()]
GoogleAds_monthly.reset_index(drop=True,inplace=True)

for i in range(len(df_op_Apex['Job No'])):
    for j in range(len(GoogleAds_monthly['jobnumber'])):
        try:
            if not (df_op_Apex['Job No'][i].startswith('000')) and (df_op_Apex['Job No'][i] in GoogleAds_monthly['jobnumber'][j]):
                print(df_op_Apex['Job No'][i])
                print(df_op_Apex['Star Date'][i])
                print(df_op_Apex['End Date'][i])
                print(df_op_Apex['APEX Budget'][i])
                print(df_op_Apex['Planner\'s Mail'][i])
                print('*' * 100)
                GoogleAds_monthly.loc[j, 'start'] = df_op_Apex['Star Date'][i]
                GoogleAds_monthly.loc[j, 'end'] = df_op_Apex['End Date'][i]
                GoogleAds_monthly.loc[j, 'budget'] = df_op_Apex['APEX Budget'][i]
                GoogleAds_monthly.loc[j, 'owner'] = df_op_Apex['Planner\'s Mail'][i]
        except Exception as e:
            pass

for k in range(len(GoogleAds_monthly['jobnumber'])):
    for l in range(len(df_normal['JOB NO'])):
        try:
            job = re.search(r"\w{4,7}\d{6}", GoogleAds_monthly['jobnumber'][k], re.I | re.M).group()
            if job in df_normal['JOB NO'][l]:
                print(job, ' is in ', df_normal['JOB NO'][l])
                print('Updating... ', df_normal['START'][l], ' to Monthly Dataframe')
                print('Updating... ', df_normal['END'][l], ' to Monthly Dataframe')
                print('Updating... ', df_normal['TOTAL MEDIA BUDGET(TWD)'][l], ' to Monthly Dataframe')
                print('Updating... ', df_normal['PLANNER MAIL'][l], ' to Monthly Dataframe')
                print('*' * 100)
                GoogleAds_monthly.loc[k, 'start'] = df_normal['START'][l]
                GoogleAds_monthly.loc[k, 'end'] = df_normal['END'][l]
                GoogleAds_monthly.loc[k, 'budget'] = df_normal['TOTAL MEDIA BUDGET(TWD)'][l]
                GoogleAds_monthly.loc[k, 'owner'] = df_normal['PLANNER MAIL'][l]
            else:
                pass
        except Exception as e:
            pass
            break

# print(GoogleAds_monthly[GoogleAds_monthly['owner'] != '']['owner'])

GoogleAds_monthly.to_excel(r"C:\Users\sebein\Desktop\結帳\Adwords\2022\2022_10\Googleads_MonthlyBilling_20221104.xlsx")
