#coding: UTF-8
import os
import shutil
import glob
import pandas as pd
import openpyxl
import xlwings as xw
import time
from pathlib import Path
import datetime


# 0.get paras from parameters.txt
paras_dict = {}
with open(r'C:\Users\035203557\OneDrive - OMRON\デスクトップ\kaizen_space\RPA\doing\parameters.txt', "r",encoding="utf-8") as f:  # encoding="utf-8" was added
    for line in f.readlines():
        line = line.strip('\n')  #remove every \n in the list

        if "=" in line :
            before_equal = line.split("=", 1)[0] # save the paras into a dict
            after_equal = line.split("=", 1)[1]  # Caution: without space in the end of the txt
            paras_dict[before_equal] = after_equal
        else:
            pass
        # print(paras_dict.items())

if "SBF_Report_path" in paras_dict:
    SBF_Report_path = paras_dict["SBF_Report_path"] # 0.1.get the path  of the SBF_Roport
else:
    print("SBF_Report_path can't be found!")

if "SBF_Report_name" in paras_dict:
    SBF_Report_name = paras_dict['SBF_Report_name'] # 0.1.get the name of the SBF_Roport
else:
    print("SBF_Report_name can't be found!")

print(SBF_Report_path)
print(SBF_Report_name)


# get the number of the sheets

    # get the path_name of the SBF_Report
SBF_Report_path_name = SBF_Report_path + "\\" + SBF_Report_name
    # open the target file

if os.path.exists(SBF_Report_path_name):
    df = pd.DataFrame

    # get all the sheets in dict
    df = pd.read_excel(SBF_Report_path_name,sheet_name=None,header=3,nrows=50,engine='openpyxl')
    # print(df['検査完了 20.05.07'])

    # get the numbers of the sheets
    num_of_sheets = len(df.keys())


    sheets = pd.DataFrame
# define the range of the data
        # change yy.mm.dd to yy/mm/dd
    for n in range(0, num_of_sheets):
        
        key_name = list(df.keys())[n]
        sheet = df[key_name]
        
        inspection_date_before_change = key_name.split(" ",1)[1]
        inspection_date_after_change = inspection_date_before_change.replace('.', '/')

        # add one column yy.mm.dd before "型式" (BX)
        sheet.rename(columns={'Unnamed: 0':'日付'}, inplace=True)
        sheet['日付'] = inspection_date_after_change
        

        # delete the column which is not necessary
        sheet = sheet.drop(columns=['検査\n不履行'])

        # delete the row which is not necessary
        sheet = sheet.dropna(axis=0, subset=["型式"])

        # save the data(range) to DataFrame
        if n == 0:
            sheets = sheet
        else:
            sheets = pd.concat([sheets,sheet], axis=0)
    sheets.rename(columns={'機能及び外観\n(Function)':'機能','化粧箱NGOK':'外観'},inplace=True)
    print(sheets)

# repeat the cycle until all the sheets has be saved to DataFrame

# wirte dataFrame to a new excel file
    with pd.ExcelWriter('path_to_file.xlsx') as writer:
        sheets.to_excel(writer)
