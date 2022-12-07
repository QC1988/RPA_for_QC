# -*- coding: UTF-8 -*-
import os
import glob
import pandas as pd
import openpyxl
import datetime
import time


# the path of folder which include the files need to be handled 
import_folder_path = os.getcwd()
export_folder_path = import_folder_path
# the path of file needs to be handled
file_path = import_folder_path + '\\' + '*.xlsx'

# all files need to be handled
files_path = glob.glob(file_path)


df = pd.DataFrame()
# concat all data vertically
# header=1 means the header is on 2nd row
for i in files_path:
    df_read_excel = pd.read_excel(i, header=6,engine='openpyxl')
    df = pd.concat([ df,df_read_excel])
# df
#axis,how,subset,inplace
df.dropna(axis=0,thresh=17,inplace=True)
i = df.shape[0]
df.dropna(axis=0, subset=['Date','OK'],inplace=True)
j = df.shape[0]

if i == j:
    print('OK.Report was created successfully.')
    today = str(datetime.date.today())
    df.to_excel(export_folder_path + '\\' + today + '.xlsx')
    time.sleep(1)

    workbook = openpyxl.load_workbook(export_folder_path + '\\'  + today + '.xlsx')
    worksheet = workbook.worksheets[0]
    worksheet.delete_cols(1)
    workbook.save(export_folder_path + '\\'  + today + '.xlsx')
else :
    print('NG.There is something wrong in the report.')
