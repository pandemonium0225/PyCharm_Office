"""刪除EXCEL裡面的圖片"""
import os
from openpyxl import load_workbook
from openpyxl import drawing
from win32com import client
images = []
for path,_,files in os.walk(r"C:\Users\sebein\Desktop\makeup schedule\202210\schedules"):
    for file in files:
        filepath = os.path.join(path,file)
        wb = load_workbook(filepath)
        wb['Daily - Broadcast']._images = []
        wb.save(filepath)