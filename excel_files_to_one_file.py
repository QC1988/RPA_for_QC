import os
import glob
import pandas as pd
import openpyxl
import datetime

# the path of folder which include the files need to be handled 
import_folder_path = os.getcwd()
export_folder_path = import_folder_path
# the path of file needs to be handled
file_path = import_folder_path + '\\' + 'book*.xlsx'
#path

# all files need to be handled
files_path = glob.glob(file_path)
#files_path

df = pd.DataFrame()
# concat all data vertically
# header=1 means the header is on 2nd row
for i in files_path:
    df_read_excel = pd.read_excel(i, header=1,engine='openpyxl')
    df = pd.concat([ df,df_read_excel])
#df

# drop column--axis=1, drop row--axis=0
df_drop = df.drop('Item3', axis=1)
df_drop
df = df_drop
#df

# ascending=True--ascending　、ascending=False--descending
df_sort = df.sort_values(by='No.', ascending=False)
df = df_sort
#df

today = str(datetime.date.today())
df.to_excel(export_folder_path + '\\' + '不良一览' + today + '.xlsx')

workbook = openpyxl.load_workbook(export_folder_path + '\\' + '不良一览' + today + '.xlsx')
worksheet = workbook.worksheets[0]
worksheet.delete_cols(1)
workbook.save(export_folder_path + '\\' + '不良一览' + today + '.xlsx')
