import pandas as pd
from datetime import datetime
import re
import numpy as np


df = pd.read_csv(r"C:\Users\sebein\Desktop\結帳\FB\test\Invoice_Report_1776662472554354 (4).csv",skiprows=9)
print(df.columns)

df_filtered = (df[
    (df.廣告主 == "PROCTER & GAMBLE TAIWAN LIMITED") |
    (df.廣告主 == "Performics of Publicis Worldwide (Hong Kong) Limited Taiwan Branch")
])

df_filtered = df_filtered[df_filtered["廣告帳號名稱"].notna()]
df_filtered["行銷活動金額"] =df_filtered['行銷活動金額'].str.replace(",","").astype("int")
df_filtered["帳單總金額"] =df_filtered['帳單總金額'].str.replace(",","").astype("int")
df_filtered['到期日'] = df_filtered['到期日'].astype('datetime64[ns]')
df_filtered['發行日期'] = df_filtered['發行日期'].astype('datetime64[ns]')
df_filtered.style.format({"到期日": lambda t: t.strftime("%y-%m-%d")})
# df_filtered['到期日'] = pd.to_datetime(df_filtered['到期日'], format='%d/%m/%Y')

# print(df_filtered["行銷活動名稱"][132])
# JobNumber = df_filtered["帳單額度說明"].str.extract(r"(PG\d{6})",flags =re.I).values.to_list()
df_filtered['Jobnumber'] = df_filtered['帳單額度說明'].str.extract(r"(PG\d{6}|SCTWFBA\d{6})")
print(df_filtered.head())

df_filtered.to_excel(r"C:\Users\sebein\Desktop\結帳\FB\test\testcredit.xlsx")