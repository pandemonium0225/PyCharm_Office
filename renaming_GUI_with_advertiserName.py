# from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
# from pdfminer.converter import TextConverter
# from pdfminer.layout import LAParams
# from pdfminer.pdfpage import PDFPage
# from io import StringIO
import os
import re
import pandas as pd
from PDF_Functions import convert_pdf_to_txt

walk = os.walk(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Aug\2022-08\Invoice")
for root, dirs, files in walk:
    for name in files:
        string = convert_pdf_to_txt(os.path.join(root,name))
        lines = list(filter(bool, string.split('\n')))
        for i in range(len(lines)):
            # "Advertiser Id" 欄位需自行轉換為數字，或在擷取TRACKER DATA時一併處理 2022/09/02 仍然需要手動調整AdvertiserID型別
            if "Advertiser Id:" in lines[i]:
                advertiser_id = lines[i].split(':',1)
                advertiser_id2 = re.search(r"\d{7,9}", advertiser_id[1])
                ref_list = pd.read_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Aug\Monthly_File_202208_original.xlsx")
                for n in range(len(ref_list)):
                    try:
                        if int(advertiser_id2.group()) == int(ref_list['AdvertiserID'][n]):
                            row_no = ref_list[ref_list.AdvertiserID == int(advertiser_id2.group())].index.tolist()
                            advertiser = ref_list['Advertiser'][row_no[0]]
                            advertiser_pattern = re.search(r"AN~(\w+\s?.*?)_MK~TW",advertiser)
                            if advertiser_pattern is None:
                                print("Renaming : {} to {}".format(name,"[" + str(row_no[0] + 2)+ "]" + name[:-4] + "_" + advertiser))
                                os.rename(os.path.join(root, name), os.path.join(root,"[" + str(row_no[0] + 2) + "]" + name[:-4] + "_" + advertiser + ".pdf"))
                                break
                            else:
                                print("Renaming : {} to {}".format(name, "[" + str(row_no[0] + 2)+ "]" + name[:-4] + "_" + advertiser_pattern.group(1)))
                                os.rename(os.path.join(root, name), os.path.join(root, "[" + str(row_no[0] + 2) + "]" + name[:-4] + "_" + advertiser_pattern.group(1) + ".pdf"))
                                break
                    except Exception as e:
                        pass
                break
