import os
from openpyxl import load_workbook
from openpyxl import drawing
from win32com import client
import win32api

def renaming_schedule():
    schedule_location = input("please paste the directory you want to rename...")
    schedule_folder = "schedules"
    folder = os.path.join(schedule_location, schedule_folder)
    os.makedirs(folder, exist_ok=True)
    for path,_,files in os.walk(schedule_location):
        for file in files:
            filepath = os.path.join(path,file)
            wb = load_workbook(filepath,data_only=True)
            sheet = wb.worksheets[3]
            img = drawing.image.Image(r"C:\Users\sebein\Desktop\makeup schedule\Sean_Stamp\sean.png")
            sheet.add_image(img,"d61")
            schedule_name = sheet.cell(10,4).value
            phase2 = schedule_name + "_P2" + ".xlsx"
            phase1 = schedule_name + "_P1" + ".xlsx"
            gam = schedule_name + "_GoogleAsia" + ".xlsx"
            client = sheet.cell(8,4).value
            # wb.save(filepath)
            if ("TW9098") in client:
                print("renaming {} to {}".format(file,phase2))
                wb.save(os.path.join(folder, phase2))
                # os.rename(filepath,os.path.join(path, phase2))
            elif ("TW8081") in client:
                print("renaming {} to {}".format(file,phase2))
                wb.save(os.path.join(folder, phase2))
                # print("renaming {} to {}".format(file,phase1))
                # wb.save(os.path.join(folder, phase1))
                # os.rename(filepath,os.path.join(path, phase1))
            elif "GOOASI" in client:
                print("renaming {} to {}".format(file,gam))
                wb.save(os.path.join(folder,gam))
            elif "TW1713" in client:
                print("renaming {} to {}".format(file, phase1))
                wb.save(os.path.join(folder,phase1))
            elif "HK9067" in client:
                print("renaming {} to {}".format(file, phase1))
                wb.save(os.path.join(folder, phase1))

            else:
                print("Non_Apex Campaign, passed")
            wb.close()


def exceltopdf(doc):
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open(doc)
    ws = wb.Worksheets[1]

    try:
        wb.SaveAs(doc[:-5], FileFormat=57)
    except Exception as e:
        print("Failed to convert")
        print(e)
    finally:
        wb.Close()
        excel.Quit()


def convert_to_pdf():
    schedule_location = input("Please input the location where the schedule excel file exists")
    for path, _, files in os.walk(schedule_location):
        for file in files:
            filepath = os.path.join(path, file)
            try:
                exceltopdf(filepath)
            except Exception as e:
                print(e)

if __name__ == '__main__':
    renaming_schedule()
    convert_to_pdf()