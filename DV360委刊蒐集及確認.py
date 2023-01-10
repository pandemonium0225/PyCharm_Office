import os
import pandas as pd
import re
import datetime
import shutil
import pygsheets

# io_directory 需注意調整，每年的FOLDER名稱可能改變
io_directory = [r'P:\PM Client\Performics\Performics Operation PG\_IO\Programmatic',
              r'P:\PM APEX Internal Practice\APEX PMP\2. Execution',
              r'P:\PM Client\Performics\Performics Operation\performance\AOD_programmatic_DV360\3 IOs',
              r'P:\PM APEX Internal Practice\APEX PMP\2. Execution\2022',
              r'P:\PM Client\Performics\PMPMP\委刊單\2022'
             ]
directory = r'D:\test' # 確認是否需要留這個變數
destination = r'D:\destination'
io_list = []
path = r'C:\Users\sebein\Desktop\結帳\DBM\2022\Dec\PG_MainFile_202212.xlsx'
data = pd.DataFrame(pd.read_excel(path))
for i in data['Insertion Order']:
    try:
        result = re.search(r"[a-zA-Z]{4,7}\d{6}",i)
        io_list.append(result.group())
    except (AttributeError,TypeError):
        pass
io_set = set(io_list)  # convert list to set將讀取到的Job Number 去重
print("本月在ref file上所出現的Job Number有...")
print("\n".join(str(job) for job in io_set))

def check_monthly_job():
    pass


# 若且惟若 Job Number "in" 讀取路徑下的檔案名稱，則COPY到指定路徑
def move_io():
    io_time = []
    for io_d in io_directory:
        for subdir, dirs, filename in os.walk(io_d):
            for file in filename:
                if file.endswith('.pdf') or file.endswith('.jpg'):
                    for io in io_set:
                        if io in file:
                            print("Congradutulation!!!,...got IO -> {}...".format(file))
                            io_file_time = (os.path.getctime(subdir + '\\' + file))
                            print(datetime.datetime.fromtimestamp(io_file_time).strftime('%Y-%m-%dT%H:%M:%S'))
                            try:
                                io_time.append({io:os.path.getmtime(subdir + '\\' + file)})
                                print("------------------------")
                                shutil.copy2(subdir + '\\' + file, destination)
                                continue
                            except Exception as e:
                                print(e)
    print(io_time)
    return io_time


def get_io_difference():
    found_io = []
    for folder, dirs, filenames in os.walk(destination):
        for file in filenames:
            try:
                job_number = re.search(r"([a-zA-Z]{4,7}\d{6,10})", file)
                found_io.append(job_number.group(1))
            except Exception as e:
                print(file,"not found")
    found_io_set = set(found_io)
    missed_io = io_set.difference(found_io_set)
    print("-----------------------")
    print("以下委刊在公槽沒有被找到")
    print("-----------------------")
    print(missed_io)
    return missed_io

def dedup_io():
    latest = {}
    job_list = []
    dedup = r"D:\destination\dedup"
    for root,subdir,filename in os.walk(destination):
        for file in filename:
            job_no = re.search(r"[a-zA-Z]{4,7}\d{6,10}", file).group()
            update_time = os.path.getmtime(root + '\\' + file)
            if job_no not in latest:
                try:
                    latest[job_no] = {file:update_time}
                    shutil.copy2(r"D:\destination\\" + file, r"D:\destination\\dedup\\" + job_no + '.pdf')
                except Exception as e:
                    print(e)
            elif job_no in latest:
                for k,v in latest[job_no].items():
                    if update_time > v:
                        updated_datetime = datetime.datetime.fromtimestamp(update_time).strftime('%Y-%m-%dT%H:%M:%S')
                        print("found updated io {}....@ {}".format(job_no,updated_datetime))
                        try:
                            latest[job_no] = {file:update_time}
                            shutil.copy2(r"D:\destination\\" + file, r"D:\destination\\dedup\\" + job_no + '.pdf')
                        except Exception as e:
                            print(e)
            else:
                print("found earlier version of IO...", file,datetime.datetime.fromtimestamp(update_time).strftime('%Y-%m-%dT%H:%M:%S'))
    print(latest)

def check_owner():
    missed_io = get_io_difference()
    gc = pygsheets.authorize(service_file=r'C:\Users\sebein\Desktop\Pythonupload\pythonupload-307303-aae13f0bea1f.json')
    sht_op = gc.open_by_url(
        # 'https://docs.google.com/spreadsheets/d/12Ts2BP9acoemXb-0bLgdBl8G6PQyMUEoUXrB0b3W8Fo/edit#gid=1390695304'
        'https://docs.google.com/spreadsheets/d/1VUNIGd9wTrKGZgMt9Fo9NWouGgrs9wBzW5aIVoZEJw4/edit#gid=1390695304'
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
    wks_apex_op = wks[1]

    df_op = wks_op.get_as_df(start="A1", numerize=False)
    df_apex = wks_apex.get_as_df(start="B1", numerize=False)

    wks_normal = sht_normal.worksheets()
    df_normal = wks_normal[0].get_as_df(start="A1", numerize=False)  # Alain 封存先改讀第五個SHEET

    # print(sht_normal.worksheets()[0].title)

    wks_pg_dv360 = sht_pg_dv360.worksheets()[0]
    df_pg_dv360 = wks_pg_dv360.get_as_df(start="A1", numerize=False)

    df_op_Apex = wks_apex_op.get_as_df(start="A1", numerize=False)

    for i in missed_io:
        for j in df_op_Apex['Job No']:
            if i in j:
                print(i)
                idx = df_op_Apex[df_op_Apex["Job No"] == j].index.values
                print("Missed 博丰委刊:" + i + "->OWNER = " + df_op_Apex["Owner"][idx].values + " ->Brand = " +
                      df_op_Apex["Client"][idx].values)

    for i in missed_io:
        for j in df_normal['JOB NO']:
            if i in j:
                print(i)
                idx = df_normal[df_normal["JOB NO"] == j].index.values
                print("Missed 博丰委刊:" + i + "->OWNER = " + df_normal["OWNER"][idx].values + " ->Brand = " +
                      df_normal["ADVERTISER"][idx].values)

    for i in missed_io:
        for j in df_apex['JobNo']:
            if i in j:
                print(i)
                idx = df_apex[df_apex["JobNo"] == j].index.values
                print("Missed 博丰委刊:" + i + "->OWNER = " + df_apex["Owner"][idx].values + " ->Brand = " +
                      df_apex["Client"][idx].values)

    for i in missed_io:
        for j in df_pg_dv360['Job']:
            if i in j:
                print(i)
                idx = df_pg_dv360[df_pg_dv360["Job"] == j].index.values
                print("Missed 博丰委刊:" + i + "->OWNER = " + df_pg_dv360["PlanningOwner"][idx].values + " ->Brand = " +
                      df_pg_dv360["Brand"][idx].values)

if __name__ == "__main__":
    move_io()
    get_io_difference()
    dedup_io()
    check_owner()