"""
TTD 整合進GOOGLE報表
"""

import pandas as pd

TTD_file_location = input(r"please input the TTD report directory with filename and extension")
DV360_file_location = input(r"please input the DV360 report directory with filename and extension")

def Merge_TTD():
    df_TTD = pd.read_excel(TTD_file_location)
    df_dv360 = pd.read_excel(DV360_file_location, skipfooter=15)
    df_TTD.rename(columns={'Advertiser ID': 'AdvertiserID',
                           'Campaign ID': 'Insertion Order ID',
                           'Advertiser Currency Code': 'Advertiser Currency',
                           'Advertiser': 'Advertiser',
                           'Campaign': 'Insertion Order',
                           'Advertiser Cost (Adv Currency)': 'MediaCost'}, inplace=True
                  )
    df_TTD.drop(df_TTD[df_TTD['MediaCost'] <= 10].index, inplace=True)
    df_TTD.sort_values(by=['Advertiser'])
    df_TTD = df_TTD.reset_index(drop=True)
    df_TTD['Partner'] = pd.Series(["TTD" for x in range(len(df_TTD.index))])
    df_TTD['Month'] = pd.Series(["2022/10" for x in range(len(df_TTD.index))])

    Merged_df = pd.concat([df_dv360, df_TTD], axis=0)
    Merged_df = Merged_df.reset_index(drop=True)
    merge_file_location = input(r"please input the file directory with filename and extension")
    Merged_df.to_excel(merge_file_location)


if __name__ == "__main__":
    Merge_TTD()

