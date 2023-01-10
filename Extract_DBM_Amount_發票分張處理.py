import pdfplumber
import pandas as pd
import re
import os
total = []
tot = {}
pair = []
t_content = []
final_df = pd.DataFrame(
    columns=[
        'order_id',
        'client_id',
        '媒體費用',
        '平台費用',
        '先前月分的無效流量調整項 (媒體費用)',
        '先前月分的無效流量調整項 (平台費用)'])

def extract_order_id(unit_invoice_content):
    adID = []
    orderIDs = re.compile(r'單(?:.*?)：.*?ID：.*?(\.00)?(\d{6,10})(?!\.00)', re.S | re.M)
    for order in orderIDs.finditer(unit_invoice_content):
#         print(order.group(2))
        adID.append(order.group(2))
#     print(adID)
    return adID

def extract_spending_items(unit_invoice_content):
    item_description = []
    spending_items = re.compile(
        r'((?<!\()平台費用(的 YouTube 預算調整項)?|媒體費用(的 YouTube 預算調整項)?|先前月分的無效流量調整項 (\(媒體費用\)|\(平台費用\))|資料費用|總費用\(平台費用 \+ 資料費用 \+ 第三方資料|第三方費用|超量放送調整項.*?(媒體費用|平台費用))',
        re.M | re.S)
    for item in spending_items.finditer(unit_invoice_content):
#         print(item.group(1))
        item_description.append(item.group(1))
#     print(item_description)
    return item_description

def extract_client_id(unit_invoice_content):
    advertiser = []
#     ad_ids = re.compile(r'戶[\s]?.*?ID：(\d{7,10})', re.S | re.M)
    ad_ids = re.compile(r'戶[\s]?.*?ID：(?:\\n)?(\d{7,10})',re.S|re.M)
    for ad_id in ad_ids.finditer(unit_invoice_content):
        advertiser.append(ad_id.group(1))
#     print(advertiser)
    return advertiser

def extract_amount_items(unit_invoice_content):
    spending = []
    amount_items = re.compile(
        r'((?<!\()平台費用(的 YouTube 預算調整項)?|媒體費用(的 YouTube 預算調整項)?|先前月分的無效流量調整項 (\(媒體費用\)|\(平台費用\))|資料費用|總費用\(平台費用 \+ 資料費用 \+ 第三方資料|第三方費用).*?(\-?\d+\.\d{2}?)',
        re.S | re.M)
    for amount_item in amount_items.finditer(unit_invoice_content):
        if amount_item.group(5) != None:
#             print(amount_item.group(5))
            spending.append(amount_item.group(5))
#     print(spending)
    return spending

def check_length(*args):
    lists = [*args]
    if all(len(lists[0]) == len(l) for l in lists[1:]):
        print(f"本張發票客戶{lists[2][0]}擷取內容成對，繼續後續操作")
        return True
    else:
        print(f"本張發票客戶{lists[2][0]}擷取內容出現問題，請查明原因")
        return False


# extract invoice info from DV360 invoice
folderpath = input("please enter the folder of Google Invoice files you want to extract?")
filepaths = [os.path.join(folderpath, name) for name in os.listdir(folderpath)]
for path in filepaths:
    with pdfplumber.open(path) as pdf:
        page_count = len(pdf.pages)
        unit_invoice_content = []
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
#                 content = re.sub(r"-","",content)
#             content = re.sub(r"\n","",content)
#                 content = re.sub(r"\s","",content)
#             print(content)
            unit_invoice_content.append(content)
        unit_invoice_content = str(unit_invoice_content)
        print(unit_invoice_content)
        print("*"*100)
        order_id = extract_order_id(unit_invoice_content)
        spending_items = extract_spending_items(unit_invoice_content)
        client_id = extract_client_id(unit_invoice_content)
        amount_items = extract_amount_items(unit_invoice_content)
        if check_length(order_id,spending_items,client_id,amount_items) == True:
            df = pd.DataFrame(columns=['order_id','spending_items','client_id','amount_items'])
            df['order_id'] = order_id
            df['spending_items'] = spending_items
            df['client_id'] = client_id
            df['amount_items'] = amount_items
            df_pivot = df.pivot(index=['client_id','order_id'],columns='spending_items',values='amount_items')
#             df_pivot = df_pivot.loc[:,['媒體費用','平台費用','先前月分的無效流量調整項 (媒體費用)','先前月分的無效流量調整項 (平台費用)']]
            df_pivot = pd.DataFrame(df_pivot.to_records())
#             print(df_pivot)
#             df_pivot.to_excel(r"C:\Users\sebein\Desktop\結帳\DBM\2022\Oct\2022-10\invoices\test_invoice\multiple_items\test.xlsx")
            final_df = pd.concat([final_df,df_pivot],ignore_index=True)

# print(final_df)
save_to_dir = input("please input the directory with file extension where the file should be saved to")
final_df.to_excel(save_to_dir)
