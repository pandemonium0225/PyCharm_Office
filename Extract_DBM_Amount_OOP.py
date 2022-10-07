def dbm_amount():
    import pdfplumber
    import pandas as pd
    import re
    import os
    total = []
    tot = {}
    pair = []
    t_content = []

    # extract invoice info from DV360 invoice
    folderpath = input("please enter the folder of Google Invoice files you want to extract?")
    filepaths = [os.path.join(folderpath, name) for name in os.listdir(folderpath)]
    for path in filepaths:
        with pdfplumber.open(path) as pdf:
            page_count = len(pdf.pages)
            adID = []
            spendItem = []
            description = []
            advertiser = []
            odd_ad = []

            for i in range(len(pdf.pages) - 2):
                content = pdf.pages[i + 2].extract_text()
                content = content.replace(",", "")
                content = content.replace("1 EA", "").replace("說說明明 金金額額 ($)", "").replace("量量 位位", "").replace("數數 單單","").replace("總金額 (TWD)", "").replace("月結單", "")
                content = re.sub(r"號碼: \d{10}", "", content)
                content = re.sub(r"Advertiser Id:\d{7}", "", content)
                content = re.sub(r"\d{4}年\d{1,2}月\d{1,2}日", "", content)
                content = re.sub(r"訂購單：\d{12}", "", content)
                content = re.sub(r"\r\n", "", content)
                content = re.sub(r"第 \d 页，共 \d 页", "", content)
                content = re.sub(r"\$\d{1,10}\.\d{2}", '', content)
                # content=re.sub(r"-","",content)
                print(content)
                t_content.append(content)

    t_content = str(t_content)
    t_content = re.sub(r"\\n", "", t_content)

    print("目前內容為\n\n" + t_content)

    # orderID=re.compile(r'單[\s]?.*?ID.*?((?<![ID~|CP~])([^\.0：]ID：(\d{7,8})(?!\.0)))',re.S|re.M)
    orderID = re.compile(r'單[\s]?.*?((?<!\_[ID~|CP~])([^\.0：]：(\d{7,8})))', re.S | re.M)
    # the "?" being removed on 2021/12/22
    # the original is r'單[\s]?.*?ID.*?((?<![ID~|CP~])([^\.0：]?\d{7,8}(?!\.0)))',re.S|re.M
    for order in orderID.finditer(t_content):
        print(order.group(3))
        adID.append(order.group(3))

    spending_items = re.compile(
        r'((?<!\()平台費用(的 YouTube 預算調整項)?|媒體費用(的 YouTube 預算調整項)?|先前月分的無效流量調整項 (\(媒體費用\)|\(平台費用\))|資料費用|總費用\(平台費用 \+ 資料費用 \+ 第三方資料|第三方費用|超量放送調整項.*?(媒體費用|平台費用))',
        re.M | re.S)
    for match in spending_items.finditer(t_content):
        print(match.group(1))
        description.append(match.group(1))

    ad_ids = re.compile(r'戶[\s]?.*?ID：(\d{7,10})', re.S | re.M)
    for ad_id in ad_ids.finditer(t_content):
        advertiser.append(ad_id.group(1))

    amount_items = re.compile(
        r'((?<!\()平台費用(的 YouTube 預算調整項)?|媒體費用(的 YouTube 預算調整項)?|先前月分的無效流量調整項 (\(媒體費用\)|\(平台費用\))|資料費用|總費用\(平台費用 \+ 資料費用 \+ 第三方資料|第三方費用).*?(\-?\d+\.\d{2}?)',
        re.S | re.M)
    for amount_item in amount_items.finditer(t_content):
        if amount_item.group(5) != None:
            print(amount_item.group(1))
            spendItem.append(amount_item.group(5))

    try:
        for i in range(len(adID)):
            expected = {adID[i]: {description[i]: spendItem[i]}}
            total.append(expected)
        for i in range(len(adID)):
            pre_pair = {adID[i]: advertiser[i]}
            pair.append(pre_pair)
    except:
        pass

    pair_final = {}
    for i in pair:
        # print(i)
        for k, v in i.items():
            print(k, v)
            if k not in pair_final:
                pair_final[k] = v

    for i in total:
        print(i)
        for k, v in i.items():
            print(k, "|", v)
            if k not in tot:
                tot[k] = v
            else:
                tot[k].update(v)

    df_dv360 = pd.DataFrame.from_dict(tot)
    df_dv360_transpose = df_dv360.T

    old_idx = df_dv360_transpose.index.to_frame()
    old_idx['advertiser id'] = old_idx[0].map(pair_final)

    old_idx

    df_dv360_transpose.index = pd.MultiIndex.from_frame(old_idx)

    print(df_dv360_transpose)

    df_dv360_transpose = df_dv360_transpose.reindex(columns=['媒體費用','平台費用','先前月分的無效流量調整項 (媒體費用)','先前月分的無效流量調整項 (平台費用)','超量放送調整項：媒體費用','超量放送調整項：平台費用','資料費用'])
    df_dv360_transpose.to_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Mar\Monthly_File_20220322_Amount.xlsx")
    # 記得EXCEL檔裡面的文字要轉換數字 就是選取之後按右鍵