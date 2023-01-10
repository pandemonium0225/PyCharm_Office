"""取出當月的Job Number 2022/2/10 改為使用read_clipboard 直接選取DATA複製後作成DATAFRAME 要記得Job 欄位需要自行對報表資料剖析"""
import os
import pandas as pd
import re
import shutil
import pygsheets
io_directory=[r'P:\PM Client\Performics\Performics Operation\performance\SEM_Google\1.SEM_starcom\3.IOs',
              r'P:\PM Client\Performics\Performics Operation\performance\SEM_Google\2.SEM_zenith\3.IOs',
              r'P:\PM APEX Internal Practice\APEX PMP\2. Execution\2021',
              r'P:\PM APEX Internal Practice\APEX PMP\2. Execution\2022',
              r'P:\PM Client\Performics\PMPMP\委刊單\2021',
              r'P:\PM Client\Performics\PMPMP\委刊單\2022',
             ]
destination = r'D:\GoogleAds_IO'
gc = pygsheets.authorize(service_file=r'C:\Users\sebein\Desktop\Pythonupload\pythonupload-307303-aae13f0bea1f.json')
GoogleAds_IO =[]
found_GoogleAdsIO=[]
# ctrl + copy google ads 後臺下載的報表，記得要把"Job"用剖析功能抓出Jobnumber資料
data = pd.read_clipboard(
    sep=",",
    header="infer",
    names=["Account name", "Customer ID", "Budget Name", "Currency Code", "Billed cost", "Job"]
)
total_result = []
job_list = data.Job
pattern = re.compile(r'((ZNTWGAD|ZNGDN|ZNTWGDN|PFTWGAD|SCTWGAD|APEX|ATYT|DGTWGAD)\d{6,9})',re.M|re.I)
for i in range(len(job_list)):
    result = pattern.findall(job_list[i])
    for j in range(len(result)):
        total_result.append(result[j][0])

total_result_set = set(total_result)
print("找到本月GOOGLEADS Campaign，計{}筆".format(len(data.Job)))
print(total_result_set)

# copy 找到的JOB NUMBER IO 到指定資料夾
for io_d in io_directory:
    for subdir,dirs,filename in os.walk(io_d):
        for f in filename:
            if (f.endswith('.pdf') or f.endswith('jpg')):
                for io in total_result_set:
                    if io in f:
                        print("Congradutulation!!!,...got IO -> {}...".format(f))
                        print(os.path.getmtime(subdir + '\\' + f))
                        GoogleAds_IO.append({io:os.path.getmtime(subdir + '\\' + f)})
                        print("---------------------")
                        shutil.copy2(subdir + '\\' + f,destination)
                        continue

for subdirs,dirs,filenames in os.walk(r"D:\GoogleAds_IO"):
    for filename in filenames:
        try:
            results=re.search(r'((ZNTWGAD|ZNGDN|ZNTWGDN|PFTWGAD|SCTWGAD|APEX|ATYT|DGTWGAD)\d{6,9})',filename)
    #         print(results.group())
            found_GoogleAdsIO.append(results.group())
        except AttributeError:
            print(filename)

found_GoogleAdsIO_set=set(found_GoogleAdsIO)
missed_GoogleAdsIO = total_result_set.difference(found_GoogleAdsIO)
print("-----------------------")
print("以下委刊在公槽沒有被找到")
print("-----------------------")
print(missed_GoogleAdsIO)

sht_op = gc.open_by_url(
# 'https://docs.google.com/spreadsheets/d/12Ts2BP9acoemXb-0bLgdBl8G6PQyMUEoUXrB0b3W8Fo/edit#gid=1390695304'
"https://docs.google.com/spreadsheets/d/1VUNIGd9wTrKGZgMt9Fo9NWouGgrs9wBzW5aIVoZEJw4/edit#gid=1390695304"
)

sht_normal = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/1J_6W-5VE55M4-hjvP0m-hwU7RFRYgRmhVehS_LG5UWM/edit#gid=1014058029'
)

sht_pg_dv360 = gc.open_by_url(
'https://docs.google.com/spreadsheets/d/1J3IXxfgy1_EUWLf_hr3vVVpdfn2zAj51sK4UUKIqK6s/edit#gid=1619667887'
)
wks = sht_op.worksheets()
wks_op=wks[1]
wks_apex = wks[0]


df_op = wks_op.get_as_df(start="A1",numerize=False)
df_apex = wks_apex.get_as_df(start="B1",numerize=False)

wks_normal = sht_normal.worksheets()
df_normal = wks_normal[0].get_as_df(start="A1",numerize=False)
df_normal.head()

wks_pg_dv360 = sht_pg_dv360.worksheets()[0]
df_pg_dv360 = wks_pg_dv360.get_as_df(start="A1",numerize=False)

for i in missed_GoogleAdsIO:
    for j in df_normal['JOB NO']:
        if i in j:
            print(i)
            idx = df_normal[df_normal["JOB NO"] == j].index.values
            print("Missed 博丰委刊:" + i + "->OWNER = " + df_normal["OWNER"][idx].values + " ->Brand = " + df_normal["ADVERTISER"][idx].values)

for i in missed_GoogleAdsIO:
    for j in df_op['Job No']:
        if i in j:
            print(i)
            idx = df_op[df_op["Job No"] == j].index.values
            print("Missed 博丰委刊:" + i + "->OWNER = " + df_op["Owner"][idx].values + " ->Brand = " + df_op["Client"][idx].values)


def googleads_dedup():
    latest = {}
    job_list = []
    dedup = r'D:\GoogleAds_IO\dedup'
    for subdirs, dirs, filenames in os.walk(r"D:\GoogleAds_IO"):
        for filename in filenames:
            try:
                job_no = re.search(r"((ZNTWGAD|ZNGDN|ZNTWGDN|PFTWGAD|SCTWGAD|APEX|ATYT|DGTWGAD)\d{6,9})", filename).group()
                update_time = os.path.getmtime(subdirs + '\\' + filename)
            except AttributeError:
                print(filename)
            for i in range(len(filenames)):
                if (job_no not in latest):
                    latest[job_no] = {filename: update_time}
                    shutil.copy2(r'D:\GoogleAds_IO' + '\\' + filename, r'D:\\GoogleAds_IO\\dedup\\' + job_no + '.pdf')
                    #                 job_list.append(filename)
                    break
                elif (job_no in latest):
                    for k, v in latest[job_no].items():
                        if update_time > v:
                            latest[job_no] = {filename: update_time}
                            #                         job_list.append(filename)
                            shutil.copy2(r'D:\GoogleAds_IO' + '\\' + filename,
                                         r'D:\\GoogleAds_IO\\dedup\\' + job_no + '.pdf')
                    break
                else:
                    pass


if __name__ == '__main__':
    googleads_dedup()