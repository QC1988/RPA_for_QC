import os
import glob
import pandas as pd
import openpyxl
import datetime
# import pyautogui

# the path of folder which include the files need to be handled 
import_folder_path = os.getcwd()
export_folder_path = import_folder_path
# the path of file needs to be handled
file_path = import_folder_path + '\\' + '*.xlsm'

# all files need to be handled
files_path = glob.glob(file_path)

df = pd.DataFrame()
# concat all data vertically
# header=1 means the header is on 2nd row
for i in files_path:
    df_read_excel = pd.read_excel(i, header=3,engine='openpyxl')
    df = pd.concat([ df,df_read_excel])

i = df.shape[0]    

df.dropna(axis=0,thresh=8,inplace=True)
j = df.shape[0]

if i == j:
    print('OK.Report was created successfully.')
    # pyautogui.alert(text="123456",title="iyoukrls",button="ok")
    today = str(datetime.date.today())
    df.to_excel(export_folder_path + '\\' + 'NGReportList' + today + '.xlsx')

    workbook = openpyxl.load_workbook(export_folder_path + '\\' + 'NGReportList' + today + '.xlsx')
    worksheet = workbook.worksheets[0]
    worksheet.delete_cols(1)
    workbook.save(export_folder_path + '\\' + 'NGReportList' + today + '.xlsx')
else:
    print('NG.There is something wrong in the report.')

