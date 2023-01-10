import os
import pandas as pd
import re
import shutil
from datetime import datetime
import pygsheets

fb_directory = [
    r'P:\PM Client\Performics\Performics Operation PG\_IO\Social',
    r'P:\PM Client\Performics\Performics Operation\performance\SEM_Google\1.SEM_starcom\3.IOs'
]

destination = r'D:\FB_IO'
io_list = []
io_list_FB=[]

def get_jobno():
    io_list_PG_FB = []
    path = input("please enter the folder of FB Reference file directory?")
    ref_df = pd.read_excel(path,sheet_name=5, skiprows=1)
    for i in range(len(ref_df)):
        _date = ref_df['Month / Year'][i]
        campaign_timestamp = pd.to_datetime(_date)
        # 需要修正跨年度時CAMPAIGN 年度/月份對不上的問題
        # if campaign_timestamp.year == (datetime.now().year-1) and campaign_timestamp.month == datetime.now().month:
        if campaign_timestamp.year == (datetime.now().year - 1) and campaign_timestamp.month == 12:
            print(ref_df['Job Number'][i])
            try:
                result = re.search(r"[a-zA-Z]{2,7}\d{6}", ref_df['Job Number'][i])
                print(f"Got PG FB Campaign's Job Number {result.group():>13} of this month...")
                io_list_PG_FB.append(result.group())
            except (AttributeError, TypeError) as e:
                print(e)
    io_set_PG_FB = set(io_list_PG_FB)
    return io_set_PG_FB


def copy_to_FB_IO():
    found_io_FB = []
    monthly_job = get_jobno()
    for io_d in fb_directory:
        for subdir,dirs,filename in os.walk(io_d):
            for f in filename:
                if (f.endswith('.pdf') or f.endswith('jpg')):
                    for io in monthly_job:
                        if io in f:
                            print("Congradutulation!!!,...got IO -> {}...".format(f))
                            print(os.path.getmtime(subdir + '\\' + f))
                            io_list_FB.append({io:os.path.getmtime(subdir + '\\' + f)})
                            print("---------------------")
                            shutil.copy2(subdir + '\\' + f,destination)
                            continue
    for subdirs, dirs, filenames in os.walk(r"D:\FB_IO"):
        for filename in filenames:
            results = re.search(r"([a-zA-Z]{2,7}\d{6})", filename)
            found_io_FB.append(results.group(1))
    found_io_FB = set(found_io_FB)
    missed_io_FB = monthly_job.difference(found_io_FB)
    print("-----------------------")
    print("以下委刊在公槽沒有被找到")
    print("-----------------------")
    print(missed_io_FB)
    return monthly_job, missed_io_FB

def check_missio_owner():
    monthly_job, missed_io_FB = copy_to_FB_IO()
    gc = pygsheets.authorize(service_file=r'C:\Users\sebein\Desktop\Pythonupload\pythonupload-307303-aae13f0bea1f.json')
    sht = gc.open_by_url(
        'https://docs.google.com/spreadsheets/d/1FFtpRXsyYooSPTB0_7e0rETVqMOK39PyLEmee-0N0IM/edit?ts=5a7d125a#gid=121235986'
    )
    sht_normal = gc.open_by_url(
        'https://docs.google.com/spreadsheets/d/1J_6W-5VE55M4-hjvP0m-hwU7RFRYgRmhVehS_LG5UWM/edit#gid=1014058029'
    )
    wks_normal = sht_normal.worksheets()
    df_normal = wks_normal[0].get_as_df(start="A1", numerize=False)
    wks_list = sht.worksheets()
    wks_tracker = wks_list[2]
    df = wks_tracker.get_as_df(start="B1", numerize=False)
    for i in missed_io_FB:
        for j in df['Job No']:
            if i == j:
                print(i)
                idx = df[df["Job No"] == i].index.values
                print("Missed 博丰委刊:" + i + "->OWNER = " + df["Owner"][idx].values + " ->Brand = " + df["Brand"][
                    idx].values)

        for k in df_normal['JOB NO']:
            if i in k:
                print(i)
                idx = df_normal[df_normal["JOB NO"] == k].index.values
                print("Missed 博丰委刊:" + i + "->OWNER = " + df_normal["OWNER"][idx].values + " ->Brand = " +
                      df_normal["ADVERTISER"][idx].values)

def dedup_and_move():
    latest_FB = {}
    job_list_FB = []
    dedup_FB = r'D:\FB_IO_Dedup'
    for subdirs, dirs, filenames in os.walk(r"D:\FB_IO"):
        for filename in filenames:
            job_no = re.search(r"[a-zA-Z]{2,7}\d{6}", filename).group()
            update_time = os.path.getmtime(subdirs + '\\' + filename)
            for i in range(len(filenames)):
                if (job_no not in latest_FB):
                    latest_FB[job_no] = {filename: update_time}
                    shutil.copy2(r'D:\FB_IO\\' + filename, r'D:\FB_IO_Dedup\\' + job_no + '.pdf')
                    break
                elif (job_no in latest_FB):
                    for k, v in latest_FB[job_no].items():
                        if update_time > v:
                            latest_FB[job_no] = {filename: update_time}
                            shutil.copy2(r'D:\FB_IO\\' + filename, r'D:\FB_IO_Dedup\\' + job_no + '.pdf')
                            break
                else:
                    pass


if __name__ == "__main__":
    check_missio_owner()
    dedup_and_move()