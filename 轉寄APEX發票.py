import win32com.client as win32
import os
import pandas as pd
import pygsheets
import pdfplumber
import re

gc = pygsheets.authorize(service_file=r'C:\Users\sebein\Desktop\Pythonupload\pythonupload-307303-aae13f0bea1f.json')
sht_APEX_PMP = gc.open_by_url(
    'https://docs.google.com/spreadsheets/d/12Ts2BP9acoemXb-0bLgdBl8G6PQyMUEoUXrB0b3W8Fo/edit#gid=1390695304')

wks = sht_APEX_PMP.worksheets()
wks_apex = wks[0]
wks_apex_op = wks[1]
df_apex = wks_apex.get_as_df(start="B1", numerize=False)
df_op_Apex = wks_apex_op.get_as_df(start="A1", numerize=False)


def send_mail_via_outlook(receiver, receiver_cc, job_no, gui_no, invoice_dirs):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = receiver
    mail.Subject = '測試寄送發票郵件' + "_" + job_no
    mail.CC = receiver_cc
    greeting_name = receiver.split(".")
    mail.Body = (
        "Hello {},\n請參考附件APEX CAMPAIGN 發票 -> JobNumber {}\n如果有問題，請與我聯絡".format(greeting_name[0],
                                                                                                job_no))
    # To attach a file to the email (optional):
    attachment = invoice_dirs
    mail.Attachments.Add(attachment)
    mail.Send()


def extract_gui_info(gui_path):
    pdf = pdfplumber.open(gui_path)
    page = pdf.pages[0]
    text = page.extract_text()
    pattern_job_no = re.compile(r"((PMSC|APEX|PMZN)\d{6,}).*?", re.M | re.S)
    pattern_bcc_no = re.compile(r"TW\d{4}\s?S?\d{4}")
    pattern_gui_no = re.compile(r"發票號碼: (\w{2}\d{8})", re.M | re.S)
    job_no = pattern_job_no.search(text)
    bcc_no = pattern_bcc_no.search(text)
    gui_no = pattern_gui_no.search(text)
    #     print(job_no.group(1),gui_no.group(1),bcc_no.group())
    return job_no.group(1), gui_no.group(1), bcc_no.group()


def get_receiverlist(job):
    receiver_list = df_apex[df_apex['JobNo'] == job]['Contact'].values[0]
    company_namepair = {'SC': "@starcomww.com", 'ZN': '@zenithmedia.com', 'PM': '@publicismedia.com',
                        'PFX': '@performics.com'}
    company = df_apex[df_apex['JobNo'] == job]['Agency'].values[0]
    receiver_list = receiver_list.split(";")
    receiver_list = [name.strip() + company_namepair[company] for name in receiver_list]
    receiver_list = [name.replace(" ", ".") for name in receiver_list]
    receiver = receiver_list[0]
    cc_list = [(';').join(receiver_list[i] for i in range(1, len(receiver_list)))]
    cc_list_final = cc_list[0]
    #     print(receiver,cc_list_final)
    return receiver, cc_list_final


def main():
    for root, dirs, filelist in os.walk(r"C:\Users\sebein\Desktop\結帳\轉寄發票\202212\test"):
        for file in filelist:
            if file.endswith('.pdf'):
                job, bcc, gui = extract_gui_info(os.path.join(root, file))
                invoice_dirs = os.path.join(root, file)
                if job.startswith('PM'):
                    receiver, cc = get_receiverlist(job)
                    print(receiver)
                    print(cc)
                    send_mail_via_outlook(receiver, cc, job, gui, invoice_dirs)
                elif job.startswith('APEX'):
                    cc = ''
                    receiver = df_op_Apex[df_op_Apex['Job No'] == job]['Planner\'s Mail']
                    if len(receiver > 0):
                        print(job, " -> ", receiver.values[0])
                        # send_mail_via_outlook(receiver.values[0],cc,job,gui,invoice_dirs)
                    else:
                        print("Something is wrong with jobnumber ", job, "...please check again or send it manually")


if __name__ == "__main__":
    main()