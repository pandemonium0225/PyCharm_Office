import pdfplumber
import pandas as pd
import re
import os

folder_path = input("Please enter the folder of DV360 invoice files to be extracted....")
file_path = [os.path.join(folder_path,name) for name in os.listdir(folder_path)]


def extract_orderID(invoice_content):
    order_id_list = []
    orderID = re.compile(r'單[\s]?.*?((?<!\_[ID~|CP~])([^\.0：]：(\d{7,10})))', re.S | re.M)
    for order in orderID.finditer(invoice_content):
        order_id_list.append(order.group(3))
    print(order_id_list)
    return order_id_list


def spending_items(invoice_content):
    description = []
    spending_description = re.compile(
    r'((?<!\()平台費用(的 YouTube 預算調整項)?|媒體費用(的 YouTube 預算調整項)?|先前月分的無效流量調整項 (\(媒體費用\)|\(平台費用\))|資料費用|總費用\(平台費用 \+ 資料費用 \+ 第三方資料|第三方費用|超量放送調整項.*?(媒體費用|平台費用))',re.M | re.S)
    for match in spending_description.finditer(invoice_content):
        description.append(match.group(1))
    print(description)
    return description


def extract_amount_items(invoice_content):
    spendItem = []
    amount_items = re.compile(r'((?<!\()平台費用(的 YouTube 預算調整項)?|媒體費用(的 YouTube 預算調整項)?|先前月分的無效流量調整項 (\(媒體費用\)|\(平台費用\))|資料費用|總費用\(平台費用 \+ 資料費用 \+ 第三方資料|第三方費用).*?(\-?\d+\.\d{2}?)',re.S | re.M)
    for amount_item in amount_items.finditer(invoice_content):
        if amount_item.group(5) is not None:
            # print(amount_item.group(1))
            spendItem.append(amount_item.group(5))
    print(spendItem)
    return spendItem


def extract_advertiser_id(invoice_content):
    advertiser_id = []
    ad_ids = re.compile(r'戶[\s]?.*?ID：(\d{7,10})', re.S | re.M)
    for ad_id in ad_ids.finditer(invoice_content):
        advertiser_id.append(ad_id.group(1))
    print(advertiser_id)
    return advertiser_id


total_content = ""
for gui in file_path:
    with pdfplumber.open(gui) as pdf:
        # total_content = []
        page_count = len(pdf.pages)
        for i in range(len(pdf.pages)-2):
            content = pdf.pages[i+2].extract_text()
            content = content.replace(",","")
            content = content.replace("1 EA","").replace("說說明明 金金額額 ($)","").replace("量量 位位","").replace("數數 單單","").replace("總金額 (TWD)","").replace("月結單","")
            content = re.sub(r"號碼: \d{10}","",content)
            content = re.sub(r"Advertiser Id:\d{7,9}","",content)
            content = re.sub(r"\d{4}年\d{1,2}月\d{1,2}日","",content)
            content = re.sub(r"訂購單：\d{12}","",content)
            content = re.sub(r"\r\n","",content)
            content = re.sub(r"第 \d 页，共 \d 页","",content)
            content = re.sub(r"\$\d{1,10}\.\d{2}",'',content)
            content = re.sub("Campaign Id:\d{7,}","",content)
            total_content += content
# total_content = re.sub(r"\n","",total_content)
# total_content = re.sub(r"\s","",total_content)
# print(total_content)
        advertiser_id = extract_advertiser_id(str(total_content))
        order_id = extract_orderID(str(total_content))
        spending_item = spending_items(str(total_content))
        amount_item = extract_amount_items(str(total_content))
        print('-------------------------file is done-----------------------------------')
#
# class gui:
#     def __init__(self,ad_id, io_id,desc,amount):
#         self.ad_id = ad_id
#         self.desc = desc
#         self.amount = amount
#         self.io_id = io_id
#         self.io_dict = {self.io_id:{desc:amount}}
#
#
# print(advertiser_id)
# print(order_id)
# print(spending_item)
# print(amount_item)
#
# case = gui(str(4176832),str(28795244),"先前月分的無效流量調整項 (平台費用)",-57)
# for i,j,k,l in zip(advertiser_id,order_id,spending_item,amount_item):
#     print(i,j,k,l)
#     print("*****************************")
#     print(case.io_dict)
#     add_item = {k:l}
#     print(add_item)
#     if k not in case.io_dict[j]:
#         case.io_dict[k] = add_item
#     else:
#         case.io_dict[j].update(add_item)
#     print(case.io_dict)
    # case = gui(i,j,k,l)
    # example = gui(i,j,k,l)

# print(example.advertiser_id)
# print(example.io_id)
# print(type(example.io_id))
# print(case.ad_id)
# print(case.io_dict)