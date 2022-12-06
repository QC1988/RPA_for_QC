# create RMA list
#!/usr/bin/python
# -*- coding: <utf-8> -*-

from calendar import month
import os
from sqlite3 import Row
from sys import path
import shutil
import glob
from numpy import pi
from openpyxl import Workbook
import pandas as pd
import openpyxl
from sqlalchemy import column, true
from sympy import N
import xlwings as xw
import time
import datetime
from pathlib import Path
import datetime
import configparser

pwd = os.getcwd()
father_path=os.path.abspath(os.path.dirname(pwd)+os.path.sep+".")
path.insert(0, father_path)

import import_paras_SBF_inspection_data

# define class to void the chars changed to lower chars
class myconf(configparser.ConfigParser):
    def __init__(self, defaults=None):
        configparser.ConfigParser.__init__(self, defaults=None)
    def optionxform(self, optionstr):
        return optionstr

Local_workpath = import_paras_SBF_inspection_data.Local_workpath

BI_source_data_online_path = import_paras_SBF_inspection_data.BI_source_data_online_path
BI_source_data_file_name = import_paras_SBF_inspection_data.BI_source_data_file_name

SBF_inspection_data_file_name = import_paras_SBF_inspection_data.SBF_inspection_data_file_name

SBF_NG_data_online_path = import_paras_SBF_inspection_data.SBF_NG_data_online_path
SBF_NG_data_file_name = import_paras_SBF_inspection_data.SBF_NG_data_file_name

SBF_inspection_data_local_path_name = import_paras_SBF_inspection_data.SBF_inspection_data_local_path_name

BI_source_data_online_path_name = import_paras_SBF_inspection_data.BI_source_data_online_path_name
BI_source_data_local_path_name = import_paras_SBF_inspection_data.BI_source_data_local_path_name

SBF_NG_data_online_path_name  = import_paras_SBF_inspection_data.SBF_NG_data_online_path_name
SBF_NG_data_local_path_name = import_paras_SBF_inspection_data.SBF_NG_data_local_path_name

# copy command
copy_command_BI_source_data_from_online_to_local =  'copy' + ' ' + '"' + BI_source_data_online_path_name  + '"'  +  ' ' + '"' + Local_workpath +'"'
copy_command_BI_source_data_from_local_to_online =  'copy' + ' ' + '"' + BI_source_data_local_path_name  + '"'  +  ' ' + '"' + BI_source_data_online_path +'"'

copy_command_SBF_NG_data_from_online_to_to_local =  'copy' + ' ' + '"' + SBF_NG_data_online_path_name  + '"'  +  ' ' + '"' + Local_workpath +'"'
copy_command_SBF_NG_data_from_local_to_online =  'copy' + ' ' + '"' + SBF_NG_data_local_path_name  + '"'  +  ' ' + '"' + SBF_NG_data_online_path +'"'

# define fuction to write xlsx, value = DataFrame type
# fuction
def write_excel_xlsx_append_xlwings(source_excel_path_name, sheet_name, last_row, dataFrame_insert_into_source, hearder_true_false):
    print("")
    print("")
    print("================================================================")
    print('    Fuction of write_excel_xlsx_append_xlwings app open.')

    print(dataFrame_insert_into_source)
    index = len(dataFrame_insert_into_source)
    print("    %d rows will be written into the file."%index)
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(source_excel_path_name)
    sheet_name = sheet_name
    sht = wb.sheets[sheet_name]
    for i in range(0, index):
        for j in range(0, len(dataFrame_insert_into_source.iloc[i,:])):
            value = dataFrame_insert_into_source.iloc[i,j]
            sht.range(i+last_row+1+hearder_true_false, j+1).value = value
    wb.save(source_excel_path_name)
    wb.close()
    app.quit()
    print('    Fuction of write_excel_xlsx_append_xlwings app closed.')
    print("    sheet %s has been written successfully."%sheet_name)
    print("================================================================")
    print("")
    print("")

# define a fuction to compare the value of two columns in source data and the data will be inserted.
def compare_two_columns_source_insert( dataFrame_source_column_list, dataFrame_insert_column_list):
    print("Checking all reports has not been import to the source data.")
    for i in dataFrame_source_column_list:
        for j in dataFrame_insert_column_list:
            if i == j:
                print("Error, %s has been import to source data or something wrong in order No."%j)
                exit()
    print("confimation OK, There is no same content in two files.")
    

def main():
    # ver.1.1 drop NA rows in SBF inspection sheet and SBF inspection NG sheet
    # DataFrame.dropna(axis=0,how='any',thresh=None,subset=None,inplace=False)
    print("================Updating the NEP inspection data.================")
    print("=       Ver.1.1                                                 =")
    print("=       2022/6/11                                               =")
    print("=       IoTG QC                                                 =")
    print("=  1.update 検査完了報告書_SBF_OSS_FY21- to file server         =")
    print("=  2.copy NG報告書 to server                                    =")
    print("=================================================================")

# 1.1 download the SBF data source
    if os.path.isfile(BI_source_data_online_path_name):
        os.system(copy_command_BI_source_data_from_online_to_local)
        print("Download the SBF source data from file server to local workpath.Please wait.")
        timedown = 5
        while timedown :
            time.sleep(1)
            if os.path.isfile(SBF_inspection_data_local_path_name):
                print("The SBF source data has been downloaded from file server to local workpath successfully.")
                break
            elif timedown >=0:
                print("Please wait.")
                timedown = timedown - 1
            elif timedown == 0:
                print("The SBF data can't be downloaded as network is abnormal.")
                exit()
    else:
        print("Error.Can't find the SBF source data.")
        exit()

# 2.1 open sheet "実績" of the BI source data with pd
    need_write_list_SBF_inspection_data_sheet = [i for i in range(13)]
    df_BI_source_data_SBF_inspection_sheet_tmp = pd.DataFrame
    df_BI_source_data_SBF_inspection_sheet_tmp = pd.read_excel(BI_source_data_local_path_name, sheet_name='実績', header=0, usecols=need_write_list_SBF_inspection_data_sheet ,engine='openpyxl')
    # ver.1.1
    df_BI_source_data_SBF_inspection_sheet_tmp.dropna(axis=0, subset=["管理No."],how='all',inplace=True)


# 2.2 open sheet "NG詳細" of the BI source data with pd
    need_write_list_SBF_NG_data_sheet = [i for i in range(8)]
    df_BI_source_data_SBF_NG_sheet_tmp = pd.DataFrame
    df_BI_source_data_SBF_NG_sheet_tmp = pd.read_excel(BI_source_data_local_path_name, sheet_name='NG詳細', header=0,usecols=need_write_list_SBF_NG_data_sheet, engine='openpyxl' )
    # ver.1.1
    df_BI_source_data_SBF_NG_sheet_tmp.dropna(axis=0, subset=["管理No."],how='all',inplace=True)

# 2.3 open sheet "実績" of the SBF inspection data with pd
    df_SBF_inspeciton_data_SBF_inspection_sheet_tmp = pd.DataFrame
    df_SBF_inspeciton_data_SBF_inspection_sheet_tmp = pd.read_excel(SBF_inspection_data_local_path_name, sheet_name='実績', header=0,usecols=need_write_list_SBF_inspection_data_sheet , engine='openpyxl')
    # print(df_SBF_inspeciton_data_SBF_inspection_sheet_tmp)
    # ver.1.1
    df_SBF_inspeciton_data_SBF_inspection_sheet_tmp.dropna(axis=0, subset=["管理No."],how='all',inplace=True)

# 2.4 open sheet "NG詳細"　of the SBF inspection data with pd
    df_SBF_inspeciton_data_SBF_NG_sheet_tmp = pd.DataFrame
    df_SBF_inspeciton_data_SBF_NG_sheet_tmp = pd.read_excel(SBF_inspection_data_local_path_name, sheet_name='NG詳細', header=0, usecols=need_write_list_SBF_NG_data_sheet, engine='openpyxl')
    # print(df_SBF_inspeciton_data_SBF_NG_sheet_tmp)
    # ver.1.1
    df_SBF_inspeciton_data_SBF_NG_sheet_tmp.dropna(axis=0, subset=["管理No."],how='all',inplace=True)

    print(df_BI_source_data_SBF_inspection_sheet_tmp)
    print(df_BI_source_data_SBF_NG_sheet_tmp)
    print(df_SBF_inspeciton_data_SBF_inspection_sheet_tmp)
    print(df_SBF_inspeciton_data_SBF_NG_sheet_tmp)
# 3.1 check sheet "実績"
    # 3.1.1 get the value(SN) of the last 30 rows in BI source data "実績" sheet
    BI_source_data_SBF_inspection_sheet_last_30_rows_KANRI_NO = df_BI_source_data_SBF_inspection_sheet_tmp.iloc[-30:, 2]
    SBF_inspection_data_SBF_inspection_sheet_divide_10 = int(df_SBF_inspeciton_data_SBF_inspection_sheet_tmp.shape[0] / 10)
    column_No = 1
    # 3.1.2 get the value(SN) of the last 1 row in SBF inspection data "実績" sheet
    for i in BI_source_data_SBF_inspection_sheet_last_30_rows_KANRI_NO:
        for j in range(1, SBF_inspection_data_SBF_inspection_sheet_divide_10):
            BI_source_data_KANRI_NO_1_row = df_SBF_inspeciton_data_SBF_inspection_sheet_tmp.iloc[-j, 2]
            if  i == BI_source_data_KANRI_NO_1_row:
                # print("j=%s"%BI_source_data_KANRI_NO_1_row)
                # print("j=%d"%int(j))
                column_No = int(j)
                break
            else:
                continue
    # solve the gap 

    column_No = column_No - 1
    if column_No == 0:
        print("No new SBF inspection data will be inserted in BI source data.")
    else:
        SBF_inspection_data_SBF_inspection_sheet_selected = df_SBF_inspeciton_data_SBF_inspection_sheet_tmp.iloc[-column_No:,:]
        # # 3.1.3 write in sheet "実績"
        last_row = df_BI_source_data_SBF_inspection_sheet_tmp.shape[0]
        write_excel_xlsx_append_xlwings(BI_source_data_local_path_name, "実績", last_row, SBF_inspection_data_SBF_inspection_sheet_selected, 1)

# 3.2 check sheet "NG詳細"
    # 3.2.1 get the value(SN) of the last 30 rows in BI source data "NG報告" sheet
    BI_source_data_NG_data_last_30_rows_KANRI_NO = df_BI_source_data_SBF_NG_sheet_tmp.iloc[-30:, 3]
    SBF_inspection_data_SBF_NG_data_sheet_divide_10 = int(df_SBF_inspeciton_data_SBF_NG_sheet_tmp.shape[0] / 10)
    checker_NO = max(50, SBF_inspection_data_SBF_NG_data_sheet_divide_10)
    column_No = 1
    # 3.2.2 get the value(SN) of the last 1 row in SBF inspection data "NG報告" sheet
    for i in BI_source_data_NG_data_last_30_rows_KANRI_NO:
        for j in range(1, checker_NO):
            BI_source_data_KANRI_NO_1_row = df_SBF_inspeciton_data_SBF_NG_sheet_tmp.iloc[-j, 3]
            if  i == BI_source_data_KANRI_NO_1_row:
                # print("j=%s"%BI_source_data_KANRI_NO_1_row)
                # print("j=%d"%int(j))
                column_No = int(j)
                break
            else:
                continue
    # solve the gap 
    column_No = column_No - 1
    if column_No == 0:
        print("No new SBF NG data will be inserted in BI source data.")
    else:
        SBF_inspection_data_SBF_NG_data_sheet_selected = df_SBF_inspeciton_data_SBF_NG_sheet_tmp.iloc[-column_No:,:]
        # # # 3.2.3 write in sheet "NG詳細"
        last_row = df_BI_source_data_SBF_NG_sheet_tmp.shape[0]
        write_excel_xlsx_append_xlwings(BI_source_data_local_path_name, "NG詳細", last_row, SBF_inspection_data_SBF_NG_data_sheet_selected, 1)


# 5 copy SBF inspection data from local to online server
    os.system(copy_command_BI_source_data_from_local_to_online)
    os.remove(BI_source_data_local_path_name)
    os.remove(SBF_inspection_data_local_path_name)

# 6 copy SBF NG reports from local to online server
    # NG reports .pdf
    reports_local_path_name_SBF_NG_data_list = []
    regEx = '*' + 'xlsx'
    reports_local_path_name_SBF_NG_data_list = glob.glob(SBF_NG_data_local_path_name + regEx, recursive=True)
    if reports_local_path_name_SBF_NG_data_list==[]:
        print("No SBF NG report will be copyed to file server.")
    else:
        print("%d SBF NG report(s) data will be copyed to file server."%len(glob.glob(SBF_NG_data_local_path_name + regEx, recursive=True)))
        for i in reports_local_path_name_SBF_NG_data_list:
            copy_command_SBF_NG_data_from_local_to_online =  'copy' + ' ' + '"' + i  + '"'  +  ' ' + '"' + SBF_NG_data_online_path +'"'
            os.system(copy_command_SBF_NG_data_from_local_to_online)
            os.remove(i)
            print("%s has been uploaded to file server."%i)
        print("All SBF NG data (.pdf) have been uploaded from local workpath to file server.")


if __name__ == '__main__':
    main()
