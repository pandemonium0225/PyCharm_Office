import os
import pandas as pd
import re
import shutil
import pygsheets
#設定搜尋委刊單的路徑
io_directory = [r'P:\PM Client\Performics\Performics Operation PG\_IO\Programmatic',
                r'P:\PM APEX Internal Practice\APEX PMP\2. Execution',
                r'P:\PM Client\Performics\Performics Operation\performance\AOD_programmatic_DV360\3 IOs',
                r'P:\PM APEX Internal Practice\APEX PMP\2. Execution',
                r'P:\PM Client\Performics\PMPMP\委刊單']

# 設定另存委刊單的路徑及其他常數變數
directory = r'D:\test'
destination = r'D:\destination'
io_list = []
found_io = []
latest = {}
job_list = []
dedup = r'D:\destination\dedup'
path = r'C:\Users\sebein\Desktop\結帳\DBM\2022\May\Monthly_File_202205.xlsx'

# 記得如果使用Github 需要另外處理憑證，不能上傳
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
wks_op = wks[2]
wks_apex = wks[0]

df_op = wks_op.get_as_df(start="A1", numerize=False)
df_apex = wks_apex.get_as_df(start="B1", numerize=False)

wks_normal = sht_normal.worksheets()
df_normal = wks_normal[0].get_as_df(start="A1", numerize=False) # Alain 封存先改讀第五個SHEET
print(sht_normal.worksheets()[0].title)

wks_pg_dv360 = sht_pg_dv360.worksheets()[0]
df_pg_dv360 = wks_pg_dv360.get_as_df(start="A1", numerize=False)

data = pd.DataFrame(pd.read_excel(path))

for i in data['Insertion Order']:
    try:
        result = re.search(r"[a-zA-Z]{4,7}\d{6}",i)
        io_list.append(result.group())
    except (AttributeError,TypeError):
        pass
io_set = set(io_list)
print("本月份結帳的JOB NUMBER為:\n",io_set)

for io_d in io_directory:
    for subdir,dirs,filename in os.walk(io_d):
        for f in filename:
            if (f.endswith('.pdf') or f.endswith('jpg')):
                for io in io_set:
                    if io in f:
                        print("Congradutulation!!!,...got IO -> {}...".format(f))
                        print(os.path.getmtime(subdir + '\\' + f))
                        try:
                            io_list.append({io:os.path.getmtime(subdir + '\\' + f)})
                            print("---------------------")
                            shutil.copy2(subdir + '\\' + f,destination)
                            continue
                        except Exception as e:
                            print(e)

for subdirs, dirs, filenames in os.walk(r"D:\destination"):
    for filename in filenames:
        results=re.search(r"([a-zA-Z]{4,7}\d{6,10})",filename)
#         print(results.group(1))
        found_io.append(results.group(1))

found_io_set=set(found_io)
missed_io = io_set.difference(found_io_set)
print("-----------------------")
print("以下委刊在公槽沒有被找到")
print("-----------------------")
print(missed_io)

for subdirs,dirs,filenames in os.walk(r"D:\destination"):
    for filename in filenames:
        job_no = re.search(r"[a-zA-Z]{4,7}\d{6,10}",filename).group()
        update_time = os.path.getmtime(subdirs + '\\' + filename)
        for i in range(len(filenames)):
            if (job_no not in latest):
                try:
                    latest[job_no]= {filename:update_time}
                    shutil.copy2(r'D:\destination' + '\\' + filename,r'D:\\destination\\dedup\\'+ job_no + '.pdf')
#                 job_list.append(filename)
                    break
                except Exception as e:
                    print(e)
            elif (job_no in latest) :
                for  k,v in latest[job_no].items():
                    if update_time > v:
                        try:
                            latest[job_no]= {filename:update_time}
#                         job_list.append(filename)
                            shutil.copy2(r'D:\destination' + '\\' + filename,r'D:\\destination\\dedup\\'+ job_no + '.pdf')
                        except Exception as e:
                            print(e)
                break
            else:
                pass

for subdirs, dirs, filenames in os.walk(r"D:\destination"):
    for filename in filenames:
        results=re.search(r"([a-zA-Z]{4,7}\d{6,10})",filename)
#         print(results.group(1))
        found_io.append(results.group(1))

found_io_set = set(found_io)
missed_io = io_set.difference(found_io_set)
print("-----------------------")
print("以下委刊在公槽沒有被找到")
print("-----------------------")
print(missed_io)

for i in missed_io:
    for j in df_normal['JOB NO']:
        if i in j:
            print(i)
            idx = df_normal[df_normal["JOB NO"] == j].index.values
            print("Missed 博丰委刊:" + i + "->OWNER = " + df_normal["OWNER"][idx].values + " ->Brand = " + df_normal["ADVERTISER"][idx].values)

for i in missed_io:
    for j in df_apex['JobNo']:
        if i in j:
            print(i)
            idx = df_apex[df_apex["JobNo"] == j].index.values
            print("Missed 博丰委刊:" + i + "->OWNER = " + df_apex["Owner"][idx].values + " ->Brand = " + df_apex["Client"][idx].values)

for i in missed_io:
    for j in df_pg_dv360['Job']:
        if i in j:
            print(i)
            idx = df_pg_dv360[df_pg_dv360["Job"] == j].index.values
            print("Missed 博丰委刊:" + i + "->OWNER = " + df_pg_dv360["PlanningOwner"][idx].values + " ->Brand = " + df_pg_dv360["Brand"][idx].values)