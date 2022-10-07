import pygsheets
import pandas as pd
import csv
import re
import datetime
from time import time

today = datetime.datetime.now()
postdate = (str(today.year) + str(today.month).zfill(2))

gc = pygsheets.authorize(service_file=r'C:\Users\sebein\Desktop\Pythonupload\pythonupload-307303-aae13f0bea1f.json')

sht_op = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/12Ts2BP9acoemXb-0bLgdBl8G6PQyMUEoUXrB0b3W8Fo/edit#gid=1390695304'
)

sht_normal = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/1J_6W-5VE55M4-hjvP0m-hwU7RFRYgRmhVehS_LG5UWM/edit#gid=1014058029'
)

sht_pg_dv360 = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/1J3IXxfgy1_EUWLf_hr3vVVpdfn2zAj51sK4UUKIqK6s/edit#gid=1619667887'
)
wks = sht_op.worksheets()
wks_op = wks[1]
wks_apex = wks[0]

df_op = wks_op.get_as_df(start="A1", numerize=False, include_tailing_empty=False)
df_apex = wks_apex.get_as_df(start="B1", numerize=False, include_tailing_empty=False)

wks_normal = sht_normal.worksheets()
df_normal = wks_normal[0].get_as_df(start="A1", numerize=False, include_tailing_empty=False)

wks_pg_dv360 = sht_pg_dv360.worksheets()[0]
df_pg_dv360 = wks_pg_dv360.get_as_df(start="A1", numerize=False, include_tailing_empty=False)

# 整理下載的DV_360 Monthly Report增加欄位
def dv360_spending_file():
    # df_monthly = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Feb\Monthly_Accrual_Report_20220301.xlsx",skipfooter=0)
    df_monthly = pd.read_excel(input("Please enter the DV360 Report File you want to compile"),skipfooter=0)
    df_monthly.rename(columns={"Advertiser ID":"AdvertiserID"},inplace=True)
    df_monthly[['MediaCost','PlatformFee','Invoice_MediaCost_InvalidRefund','Invoice_PlatformFee_InvalidRefundInvoice_Total','TrueView_Adjust_media','TrueView_Adjust_platform','ThridParty','CM','Perfomics_Invoice','Agency_Note','PFX_Act','Startdate','EndDate','Contact','Budget']] = ''
    df_monthly['AdvertiserID'] = df_monthly['AdvertiserID'].astype('str')
    df_monthly['Insertion Order'] = df_monthly['Insertion Order'].fillna('').apply(str)
    return df_monthly

def extract_job(df_monthly):
    regex_list = [
        # r"(?:PGTW\d{6}.*?@)",
        r"(?:PGTW\d{6}.* ?CN~[0-9A-Za-z\s\ / \-\(\)\]]*(?! >: @))",
        r"(?:PFTW.*?\d{6}.*?])",
        r"(?:APEX|PMSC|PMZN|PMTW)\d{6}_.*?_",
        r"(?:ZNTWGDV|SCTWGDV|PMTWGDV)\d{6}.*?"
    ]
    df_monthly['BCC_KeyIn'] = df_monthly['Insertion Order'].str.extract("(" + "|".join(regex_list) + ")") + "_" + postdate
    return df_monthly


def timer_func(func):
    # This function shows the execution time of
    # the function object passed
    def wrap_func(*args, **kwargs):
        t1 = time()
        result = func(*args, **kwargs)
        t2 = time()
        print(f'Function {func.__name__!r} executed in {(t2 - t1):.4f}s')
        return result

    return wrap_func


@timer_func
def op_apex(df_monthly):
    print("Starting op_apex \n")
    for i in range(len(df_op['Job No'])):
        for j in range(len(df_monthly['Insertion Order'])):
            try:
                if not (df_op['Job No'][i].startswith('000')) and (
                        df_op['Job No'][i] in df_monthly['Insertion Order'][j]):
                    print(df_op['Job No'][i])
                    print(df_op['Star Date'][i])
                    print(df_op['End Date'][i])
                    print(df_op['APEX Budget'][i])
                    print(df_op['Planner\'s Mail'][i])
                    print('*' * 100)
                    df_monthly.loc[j, 'Startdate'] = df_op['Star Date'][i]
                    df_monthly.loc[j, 'EndDate'] = df_op['End Date'][i]
                    df_monthly.loc[j, 'Budget'] = df_op['APEX Budget'][i]
                    df_monthly.loc[j, 'Contact'] = df_op['Planner\'s Mail'][i]
                    break
            except:
                pass
    print("op_apex part completed...")


# function pmp_apex works
@timer_func
def pmp_apex(df_monthly):
    print("Initiating...function'pmp_apex'\n")
    for i in range(len(df_apex['JobNo'])):
        for k in range(len(df_monthly['Insertion Order'])):
            try:
                if (df_apex['JobNo'][i] != '') and (df_apex['JobNo'][i] in df_monthly['Insertion Order'][k]):
                    print(df_apex['JobNo'][i])
                    print(df_apex['Startdate'][i])
                    print(df_apex['EndDate'][i])
                    print(df_apex['Budget'][i])
                    print(df_apex['Contact'][i])
                    print('*' * 100)
                    df_monthly.loc[k, 'Startdate'] = df_apex['Startdate'][i]
                    df_monthly.loc[k, 'EndDate'] = df_apex['EndDate'][i]
                    df_monthly.loc[k, 'Budget'] = df_apex['Budget'][i]
                    df_monthly.loc[k, 'Contact'] = df_apex['Contact'][i]
                    break
            except:
                print('Wrong...', df_apex['JobNo'][i], df_monthly['Insertion Order'][k])
                break
    print("function'pmp_apex' is done...\n")


# function op_normal
@timer_func
def op_normal(df_monthly):
    for i in range(len(df_monthly['Insertion Order'])):
        for l in range(len(df_normal['JOB NO'])):
            try:
                job = re.search(r"\w{4,7}\d{6}", df_monthly['Insertion Order'][i], re.I | re.M).group()
                if job in df_normal['JOB NO'][l]:
                    print(job, ' is in ', df_normal['JOB NO'][l])
                    print('Updating... ', df_normal['START'][l], ' to Monthly Dataframe')
                    print('Updating... ', df_normal['END'][l], ' to Monthly Dataframe')
                    print('Updating... ', df_normal['TOTAL MEDIA BUDGET(TWD)'][l], ' to Monthly Dataframe')
                    print('Updating... ', df_normal['PLANNER MAIL'][l], ' to Monthly Dataframe')
                    print('*' * 100)
                    df_monthly.loc[i, 'Startdate'] = df_normal['START'][l]
                    df_monthly.loc[i, 'EndDate'] = df_normal['END'][l]
                    df_monthly.loc[i, 'Budget'] = df_normal['TOTAL MEDIA BUDGET(TWD)'][l]
                    df_monthly.loc[i, 'Contact'] = df_normal['PLANNER MAIL'][l]
                    break
                else:
                    pass
            except:
                break


# function pg_dv360
@timer_func
def pg_dv360(df_monthly):
    for i in range(len(df_monthly['Insertion Order'])):
        for k in range(len(df_pg_dv360['Job'])):
            try:
                pg_job = re.search(r"\w{4,7}\d{6}", df_monthly['Insertion Order'][i], re.I | re.M).group()
                if pg_job in df_pg_dv360['Job'][k]:
                    print(pg_job, ' is in ', df_pg_dv360['Job'][k])
                    print('updating...,', df_pg_dv360['Start Date'][k])
                    print('updating...,', df_pg_dv360['End Date'][k])
                    print('updating...,', df_pg_dv360['Budget(Local Currency)'][k])
                    print('updating...,', df_pg_dv360['PlanningOwner'][k])
                    print('*' * 100)
                    df_monthly.loc[i, 'Startdate'] = df_pg_dv360['Start Date'][k]
                    df_monthly.loc[i, 'EndDate'] = df_pg_dv360['End Date'][k]
                    df_monthly.loc[i, 'Budget'] = df_pg_dv360['Budget(Local Currency)'][k]
                    df_monthly.loc[i, 'Contact'] = df_pg_dv360['PlanningOwner'][k]
                    break
            except:
                pass

if __name__ == '__main__':
    df_monthly = dv360_spending_file()
    df_monthly = extract_job(df_monthly)
    op_normal(df_monthly)
    pmp_apex(df_monthly)
    op_apex(df_monthly)
    pg_dv360(df_monthly)
    df_monthly.to_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Sep\\" + "Monthly_File_" + postdate + "pycharm_test" + ".xlsx",index=False)

