import os
import pandas as pd
import re
import shutil

fb_directory = [
    r'P:\PM Client\Performics\Performics Operation PG\_IO\Social',
    r'P:\PM Client\Performics\Performics Operation\performance\SEM_Google\1.SEM_starcom\3.IOs'
]

destination = r'D:\FB_IO'
io_list = []

def get_jobno():
    io_list_PG_FB = []
    path = input("please enter the folder of FB Reference file directory?")
    ref_df = pd.read_excel(path)
    for i in ref_df['Job Number']:
        try:
            result = re.search(r"[a-zA-Z]{2,7}\d{6}", i)
            print(f"Got PG FB Campaign's Job Number {result.group():>13} of this month...")
            io_list_PG_FB.append(result.group())
        except (AttributeError, TypeError) as e:
            print(e)
            pass
    io_set_PG_FB = set(io_list_PG_FB)
    return io_set_PG_FB







if __name__ == "__main__":
    get_jobno()
